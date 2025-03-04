//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_IsolPad_99_PE34.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_IsolPad_99_PE34
//   Author       :  Hema
//   Creation Date:  1.August.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1.August.2012    Hema     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   23.Nov.2012     Hema      CR-CP-219113 Modified the code with SymbolGeomHelper
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Bline_IsolPad_99_PE34 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_IsolPad_99_PE34"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "GA", "Gap between Pad and Web Rail", 0)]
        public InputDouble m_dGA;
        [InputDouble(3, "LC", "Length of Plate", 0)]
        public InputDouble m_dLC;
        [InputDouble(4, "LS", "Distance from edge to hole", 0)]
        public InputDouble m_dLS;
        [InputDouble(5, "WC", "Width of Plate", 0)]
        public InputDouble m_dWC;
        [InputDouble(6, "TH", "Thickness of Plate", 0)]
        public InputDouble m_dTH;
        [InputDouble(7, "TrayWT", "Tray Web Thickness", 0)]
        public InputDouble m_dTrayWT;
        [InputDouble(8, "Inside_Outside", "Inside or Outside", 1)]
        public InputDouble m_dInside_Outside;
        [InputDouble(9, "Clamp_Guide", "Clamp or Guide", 1)]
        public InputDouble m_dClamp_Guide;
        [InputDouble(10, "WT", "Tray Width", 0)]
        public InputDouble m_dWT;

        #endregion

        #region "Definitions of Aspects and their outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("L_ISOLATOR", "L_ISOLATOR")]
        [SymbolOutput("R_ISOLATOR", "R_ISOLATOR")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                
                double GA = m_dGA.Value;
                double LC = m_dLC.Value;
                double WC = m_dWC.Value;
                double TH = m_dTH.Value;
                double TW = m_dTrayWT.Value;
                double inOrOut = m_dInside_Outside.Value;
                double WT = m_dWT.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                if (LC <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidPlateL, "Plate Length value should  be greater than 0."));
                if (WC <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidPlateW, "Plate Width  value should  be greater than 0."));
                if (TH <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidPlateT, "Plate Thickness value should  be greater than 0."));

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                PropertyValueCodelist insideOutsideCodelistValue = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                if (inOrOut <= 0 || inOrOut > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInorOut, "inOrOut value should be between 1 and 2"));

                if (insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem((int)inOrOut).DisplayName.ToLower() == "outside")
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW + GA, TH / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                    Projection3d leftIsolator = (Projection3d)symbolGeometryHelper.CreateBox(null, LC, WC, TH);
                    m_Symbolic.Outputs["L_ISOLATOR"] = leftIsolator;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 + TW + GA), TH / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                    Projection3d rightIsolator = (Projection3d)symbolGeometryHelper.CreateBox(null, LC, WC, TH);
                    m_Symbolic.Outputs["R_ISOLATOR"] = rightIsolator;
                }
                else
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 - GA - LC, TH / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                    Projection3d leftIsolator = (Projection3d)symbolGeometryHelper.CreateBox(null, LC, WC, TH);
                    m_Symbolic.Outputs["L_ISOLATOR"] = leftIsolator;

                    symbolGeometryHelper.ActivePosition = new Position(0, (-WT / 2 + GA + LC), TH / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                    Projection3d rightIsolator = (Projection3d)symbolGeometryHelper.CreateBox(null, LC, WC, TH);
                    m_Symbolic.Outputs["R_ISOLATOR"] = rightIsolator;
                }
            }
            catch  //General Unhandled exception 
            {
               
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_IsolPad_99_PE34."));
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
                PropertyValueCodelist insideOutsideCodelistValue, clampGuideCodelistValue;

                insideOutsideCodelistValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                if (insideOutsideCodelistValue.PropValue <= 0 || insideOutsideCodelistValue.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInside_Outside, "Bline Inside_Outside Code list value should be between 1 and 2"));
                string inside_Outside = insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem(insideOutsideCodelistValue.PropValue).DisplayName;

                clampGuideCodelistValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineClampGuide", "Clamp_Guide");
                if (clampGuideCodelistValue.PropValue <= 0 || clampGuideCodelistValue.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidClampGuide, "Clamp Guide Code list value should be between 1 and 2"));
                string clamp_Guide = clampGuideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem(clampGuideCodelistValue.PropValue).DisplayName;
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                bomDescription = catalogPart.PartDescription + ", Installed: " + inside_Outside + ", Underneath :" + clamp_Guide;
                return bomDescription;
            }
            catch 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_IsolPad_99_PE34."));
                return "";
            }
        }
        #endregion
    }
}





