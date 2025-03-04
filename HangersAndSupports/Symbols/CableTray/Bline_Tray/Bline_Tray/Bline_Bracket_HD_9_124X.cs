//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_Bracket_HD_9_124X.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Bracket_HD_9_124X
//   Author       :  Hema
//   Creation Date:  9.August.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   9.August.2012   Hema       CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   23.Nov.2012     Rajeswari  CR-CP-219113 Modified the code with SymbolGeomHelper
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
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
    public class Bline_Bracket_HD_9_124X : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Bracket_HD_9_124X"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "WP", "Width of Bracket", 0)]
        public InputDouble m_dWP;
        [InputDouble(3, "LC", "Length of Bracket", 0)]
        public InputDouble m_dLC;
        [InputDouble(4, "LS", "Distance from edge to hole", 0)]
        public InputDouble m_dLS;
        [InputDouble(5, "LP", "Length of Bottom Plate", 0)]
        public InputDouble m_dLP;
        [InputDouble(6, "HC", "Height of Bracket", 0)]
        public InputDouble m_dHC;
        [InputDouble(7, "TH", "Thickness of Bracket", 0)]
        public InputDouble m_dTH;
        [InputDouble(8, "DO", "Hole Diameter", 0)]
        public InputDouble m_dDO;
        [InputDouble(9, "TrayWT", "Hole Diameter", 0)]
        public InputDouble m_dTrayWT;
        [InputDouble(10, "Gap", "Hole Diameter", 0)]
        public InputDouble m_dGap;
        [InputDouble(11, "WT", "Hole Diameter", 0)]
        public InputDouble m_dWT;
        [InputDouble(12, "NumBolts", "Number of Bolts", 2)]
        public InputDouble m_dNumBolts;
        [InputDouble(13, "LO", "Distance between bolts", 0)]
        public InputDouble m_dLO;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("L_BRACKET", "L_BRACKET")]
        [SymbolOutput("R_BRACKET", "R_BRACKET")]
        [SymbolOutput("L_TRAYBOLT1", "L_TRAYBOLT1")]
        [SymbolOutput("L_TRAYBOLT2", "L_TRAYBOLT2")]
        [SymbolOutput("R_TRAYBOLT1", "R_TRAYBOLT1")]
        [SymbolOutput("R_TRAYBOLT2", "R_TRAYBOLT2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                double WP = m_dWP.Value;
                double LC = m_dLC.Value;
                double LS = m_dLS.Value;
                double LP = m_dLP.Value;
                double HC = m_dHC.Value;
                double TH = m_dTH.Value;
                double DO = m_dDO.Value;
                double TW = m_dTrayWT.Value;
                double WT = m_dWT.Value;
                int numBolts = (int)m_dNumBolts.Value;
                double LO = m_dLO.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                if (DO <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDO, "Hole Diameter value should  be greater than 0."));
                //ports
                Port trayPort = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = trayPort;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW, HC));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW, HC - LP));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW + LC - LP, 0));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW + LC, 0));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW + LC, TH));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW + LC - LP, TH));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW + TH, HC - LP));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW + TH, HC));
                pointCollection.Add(new Position(WP / 2, WT / 2 + TW, HC));

                Vector projVec = new Vector(-WP, 0, 0);
                Projection3d lBracket = new Projection3d(OccurrenceConnection, new LineString3d(pointCollection), projVec, projVec.Length, true);
                m_PhysicalAspect.Outputs["L_BRACKET"] = lBracket;
                pointCollection.Clear();

                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW), HC));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW), HC - LP));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + LC - LP), 0));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + LC), 0));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + LC), TH));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + LC - LP), TH));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + TH), HC - LP));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + TH), HC));
                pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW), HC));

                projVec.Set(-WP, 0, 0);
                Projection3d rBracket = new Projection3d(OccurrenceConnection, new LineString3d(pointCollection), projVec, projVec.Length, true);
                m_PhysicalAspect.Outputs["R_BRACKET"] = rBracket;

                if (numBolts == 2)
                {
                    Vector normal = new Position(0, WT / 2 + TW + 1.25 * TH, HC - LS).Subtract(new Position(0, WT / 2 + TW, HC - LS));
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW, HC - LS);
                    symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                    Projection3d lTrayBolt1 = symbolGeometryHelper.CreateCylinder(null, DO / 2, normal.Length);
                    m_PhysicalAspect.Outputs["L_TRAYBOLT1"] = lTrayBolt1;

                    normal = new Position(0, -(WT / 2 + TW + 1.25 * TH), HC - LS).Subtract(new Position(0, -(WT / 2 + TW), HC - LS));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 + TW), HC - LS);
                    symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                    Projection3d rTrayBolt1 = symbolGeometryHelper.CreateCylinder(null, DO / 2, normal.Length);
                    m_PhysicalAspect.Outputs["R_TRAYBOLT1"] = rTrayBolt1;

                    Port rodHole1 = new Port(OccurrenceConnection, part, "RodHole1", new Position(0, WT / 2 + TW + LC - LS, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port2"] = rodHole1;

                    Port rodHole2 = new Port(OccurrenceConnection, part, "RodHole2", new Position(0, -(WT / 2 + TW + LC - LS), 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port3"] = rodHole2;
                }
                else if (numBolts == 4)
                {
                    Vector normal = new Position(-LO / 2, WT / 2 + TW + 1.25 * TH, HC - LS).Subtract(new Position(-LO / 2, WT / 2 + TW, HC - LS));
                    symbolGeometryHelper.ActivePosition = new Position(-LO / 2, WT / 2 + TW, HC - LS);
                    symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                    Projection3d lTrayBolt1 = symbolGeometryHelper.CreateCylinder(null, DO / 2, normal.Length);
                    m_PhysicalAspect.Outputs["L_TRAYBOLT1"] = lTrayBolt1;

                    normal = new Position(LO / 2, WT / 2 + TW + 1.25 * TH, HC - LS).Subtract(new Position(LO / 2, WT / 2 + TW, HC - LS));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(LO / 2, WT / 2 + TW, HC - LS);
                    symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                    Projection3d lTrayBolt2 = symbolGeometryHelper.CreateCylinder(null, DO / 2, normal.Length);
                    m_PhysicalAspect.Outputs["L_TRAYBOLT2"] = lTrayBolt2;

                    normal = new Position(-LO / 2, -(WT / 2 + TW + 1.25 * TH), HC - LS).Subtract(new Position(-LO / 2, -(WT / 2 + TW), HC - LS));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-LO / 2, -(WT / 2 + TW), HC - LS);
                    symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                    Projection3d rTrayBolt1 = symbolGeometryHelper.CreateCylinder(null, DO / 2, normal.Length);
                    m_PhysicalAspect.Outputs["R_TRAYBOLT1"] = rTrayBolt1;

                    normal = new Position(LO / 2, -(WT / 2 + TW + 1.25 * TH), HC - LS).Subtract(new Position(LO / 2, -(WT / 2 + TW), HC - LS));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(LO / 2, -(WT / 2 + TW), HC - LS);
                    symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                    Projection3d rTrayBolt2 = symbolGeometryHelper.CreateCylinder(null, DO / 2, normal.Length);
                    m_PhysicalAspect.Outputs["R_TRAYBOLT2"] = rTrayBolt2;

                    Port rodHole1 = new Port(OccurrenceConnection, part, "RodHole1", new Position(-LO / 2, WT / 2 + TW + LC - LS, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port2"] = rodHole1;

                    Port rodHole2 = new Port(OccurrenceConnection, part, "RodHole2", new Position(-LO / 2, -(WT / 2 + TW + LC - LS), 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port3"] = rodHole2;

                    Port rodHole3 = new Port(OccurrenceConnection, part, "RodHole3", new Position(LO / 2, WT / 2 + TW + LC - LS, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port4"] = rodHole3;

                    Port rodHole4 = new Port(OccurrenceConnection, part, "RodHole4", new Position(LO / 2, -(WT / 2 + TW + LC - LS), 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port5"] = rodHole4;
                }
            }
            catch
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_Bracket_HD_9_124X."));
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
                PropertyValueCodelist materialCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineBracketMat", "Material");
                if (materialCodelist.PropValue <= 0 || materialCodelist.PropValue > 3)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBracketMaterial, "Bracket Material Code list value should be between 1 and 3"));
                    return "";
                }

                string materIal = materialCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(materialCodelist.PropValue).DisplayName;
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                bomDescription = catalogPart.PartDescription + ", Material: " + materIal;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_Bracket_HD_9_124X."));
                return "";
            }
        }

        #endregion
    }
}



