//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_Clamp_9_120X.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Clamp_9_120X
//   Author       :  Shilpi	
//   Creation Date:  08.August.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   08.August.2012  Shilpi     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   23.Nov.2012     Rajeswari  CR-CP-219113 Modified the code with SymbolGeomHelper
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//   07.Sep.2015     PR         TR 277225	B-Line Hangers do not place correctly
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
    public class Bline_Clamp_9_120X : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Clamp_9_120X"
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
        [InputDouble(9, "Inside_Outside", "Inside or Outside", 2.0)]
        public InputDouble m_dInside_Outside;
        [InputDouble(10, "WT", "Tray Width", 0)]
        public InputDouble m_dWT;
        [InputString(11, "WithHardware", "With or Without Hardware", "")]
        public InputString m_strWithHardware;
        [InputDouble(12, "Clamp_Guide", "Clamp or Guide", 1)]
        public InputDouble m_dClamp_Guide;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("L_CLAMP", "L_CLAMP")]
        [SymbolOutput("R_CLAMP", "R_CLAMP")]
        [SymbolOutput("L_INTHRDRH", "L_INTHRDRH")]
        [SymbolOutput("R_INTHRDRH", "R_INTHRDRH")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]

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
                double inOrOut =m_dInside_Outside.Value;
                double WT = m_dWT.Value;
                string withHardware = m_strWithHardware.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                PropertyValueCodelist insideOutsideCodelistValue = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                if (inOrOut <= 0 || inOrOut > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInorOut, "inOrOut value should be between 1 and 2"));
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

                //ports
                Port trayPort = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = trayPort;

                if (insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem((int)inOrOut).DisplayName.ToLower() == "outside")
                {
                    if (withHardware.ToLower() == "without")
                    {
                        Port rodHole2 = new Port(OccurrenceConnection, part, "RodHole2", new Position(0, 0.5 * (WT + TW) + 0.5 * TW + GA + LS, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Port2"] = rodHole2;

                        Port rodHole1 = new Port(OccurrenceConnection, part, "RodHole1", new Position(0, -(0.5 * (WT + TW) + 0.5 * TW + GA + LS), 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Port3"] = rodHole1;
                    }
                    else
                    {
                        Vector normal1 = new Position(0, 0.5 * (WT + TW)  + LS, HC * 1.25).Subtract(new Position(0, 0.5 * (WT + TW)  + LS, 0));
                        symbolGeometryHelper.ActivePosition = new Position(0, 0.5 * (WT + TW)  + LS, 0);
                        symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                        Projection3d lInthrfrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                        m_PhysicalAspect.Outputs["L_INTHRDRH"] = lInthrfrh;

                        Vector normal2 = new Position(0, -(0.5 * (WT + TW)  + LS), HC * 1.25).Subtract(new Position(0, -(0.5 * (WT + TW) + LS), 0));
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, -(0.5 * (WT + TW)  + LS), 0);
                        symbolGeometryHelper.SetOrientation(normal2, new Vector(1, 0, 0));
                        Projection3d rInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal2.Length);
                        m_PhysicalAspect.Outputs["R_INTHRDRH"] = rInthrdrh;
                    }

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2+TW +CL/2, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d lClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, HC, WC, CL);
                    m_PhysicalAspect.Outputs["L_CLAMP"] = lClamp;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 + TW + CL / 2), 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d rClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, HC, WC, CL);
                    m_PhysicalAspect.Outputs["R_CLAMP"] = rClamp;

                }
                else
                {
                    if (withHardware.ToLower() == "without")
                    {
                        Port rodHole2 = new Port(OccurrenceConnection, part, "RodHole2", new Position(0, 0.5 * WT - GA - LS, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["RodHole2"] = rodHole2;

                        Port rodHole1 = new Port(OccurrenceConnection, part, "RodHole1", new Position(0, -0.5 * WT + GA + LS, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["RodHole1"] = rodHole1;
                    }
                    else
                    {
                        Vector normal1 = new Position(0, 0.5 * WT - GA - LS, HC * 1.25).Subtract(new Position(0, 0.5 * WT - GA - LS, 0));
                        symbolGeometryHelper.ActivePosition = new Position(0, 0.5 * WT - GA - LS, 0);
                        symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                        Projection3d lInthrfrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                        m_PhysicalAspect.Outputs["L_INTHRDRH"] = lInthrfrh;

                        Vector normal2 = new Position(0, -0.5 * WT + GA + LS, HC * 1.25).Subtract(new Position(0, -0.5 * WT + GA + LS, 0));
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, -0.5 * WT + GA + LS, 0);
                        symbolGeometryHelper.SetOrientation(normal2, new Vector(1, 0, 0));
                        Projection3d rInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal2.Length);
                        m_PhysicalAspect.Outputs["R_INTHRDRH"] = rInthrdrh;
                    }
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
            }


            catch 
            {
                
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_Clamp_9_120X."));
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
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)part.GetPropertyValue("IJUAHgrBlineFinish", "Finish");
                PropertyValueCodelist insideOutsideCodelistValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                PropertyValueCodelist ClampOrGuideCodelistValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineClampGuide", "Clamp_Guide");
                if (finishCodelist.PropValue < 0 || finishCodelist.PropValue > 8)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBlineFinish, "Bline Finish Code list value should be between 1 and 8"));
                if (insideOutsideCodelistValue.PropValue <= 0 || insideOutsideCodelistValue.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInside_Outside, "Bline Inside_Outside Code list value should be between 1 and 2"));
                if (ClampOrGuideCodelistValue.PropValue <= 0 || ClampOrGuideCodelistValue.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidClampGuide, "Clamp Guide Code list value should be between 1 and 2"));

                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                string inOrOut = insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem(insideOutsideCodelistValue.PropValue).DisplayName;
                string clampOrGuide = ClampOrGuideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem(ClampOrGuideCodelistValue.PropValue).DisplayName;


                bomDescription = part.PartDescription + "," + clampOrGuide + ",Installed " + inOrOut + ", Finish: " + finish;

                return bomDescription;
            }

            catch 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_Clamp_9_120X."));
                return "";
               
            }
        }

        #endregion
    }
}



