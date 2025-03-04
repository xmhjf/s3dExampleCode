//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_Beam_Clamp_B750.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Beam_Clamp_B750
//   Author       :  Vijaya
//   Creation Date:  30.July.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.July.2012    Vijaya     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   26.Nov.2012     Rajeswari  CR-CP-219113 Modified the code with SymbolGeomHelper
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
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
    public class Bline_Beam_Clamp_B750 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbolss.Bline_Beam_Clamp_B750"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ClampLength", "Length of the Clamp", 0)]
        public InputDouble m_dClampLength;
        [InputDouble(3, "ClampWidth", "Width of the Clamp", 0)]
        public InputDouble m_dClampWidth;
        [InputDouble(4, "TH", "Thickness of Clamp", 0)]
        public InputDouble m_dThickness;
        [InputDouble(5, "DistEdge", "Distance from edge to hole", 0)]
        public InputDouble m_dDistEdge;
        [InputDouble(6, "HoleDiameter", "Hole Diameter", 0)]
        public InputDouble m_dHoleDiameter;
        [InputDouble(7, "WF", "Flange Width", 0)]
        public InputDouble m_dWF;
        [InputDouble(8, "MinWF", "Minimum Flange Width", 0)]
        public InputDouble m_dMinWF;
        [InputDouble(9, "MaxWF", "Maximum Flange Width", 0)]
        public InputDouble m_dMaxWF;
        [InputDouble(10, "FT", "Flange Thickness", 0)]
        public InputDouble m_dFT;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CLAMP_BODY", "CLAMP_BODY")]
        [SymbolOutput("JHOOK_B", "JHOOK_B")]
        [SymbolOutput("JHOOK_M", "JHOOK_M")]
        [SymbolOutput("JHOOK_T", "JHOOK_T")]

        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;


                double CL = m_dClampLength.Value;
                double WC = m_dClampWidth.Value;
                double TH = m_dThickness.Value;
                double LS = m_dDistEdge.Value;
                double DO = m_dHoleDiameter.Value;
                double WF = m_dWF.Value;
                double minWF = m_dMinWF.Value;
                double maxWF = m_dMaxWF.Value;
                double FT = m_dFT.Value;

                //Exceptions
                if (FT <= 0.0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidFT, "Flange Thickness should be greater than 0"));
                    return;
                }
                if (WF <= 0.0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidWF, "Flange Width should be greater than 0"));
                    return;
                }
                if (WF > maxWF || WF < minWF)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidWFRange, "Flange Width should be between Minimum and Maximum Value"));

                }
                //=================================================
                //Construction of Physical Aspect    
                //=================================================
                if (DO <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDO, "Hole Diameter value should  be greater than 0."));
                }
                if (CL <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidCL, "Clamp Length value should  be greater than 0."));
                }
                if (WC <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidWCgtZero, "Width of Clamp value should  be greater than 0."));
                }
                if (LS < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidLS, "Distance from edge to hole value should not be less than 0."));
                }

                //Add Ports
                Port port1 = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "InThrdRH1", new Position(0, -WC / 2, 1.5 * FT + TH + 0.008), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -DO / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d jhookB = (Projection3d)symbolGeometryHelper.CreateBox(null, WF, DO, DO);
                m_PhysicalAspect.Outputs["JHOOK_B"] = jhookB;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, WF + DO / 2, -DO);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d jhookM = (Projection3d)symbolGeometryHelper.CreateBox(null, DO + FT, DO, DO);
                m_PhysicalAspect.Outputs["JHOOK_M"] = jhookM;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, WF + DO, FT + DO / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                Projection3d jhookT = (Projection3d)symbolGeometryHelper.CreateBox(null, 2 * DO, DO, DO);
                m_PhysicalAspect.Outputs["JHOOK_T"] = jhookT;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -WC / 2, -(LS));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d clampBody = (Projection3d)symbolGeometryHelper.CreateBox(null, (TH + 1.5 * FT + LS), CL, WC);
                m_PhysicalAspect.Outputs["CLAMP_BODY"] = clampBody;
            }
            catch
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_Beam_Clamp_B750."));
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
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAHgrBlineFinish", "Finish");
                if (finishCodelist.PropValue < 0 || finishCodelist.PropValue > 8)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBlineFinish, "Bline Finish Code list value should be between 1 and 8"));
                }

                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                bomDescription = catalogPart.PartDescription + ", Finish: " + finish;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_Beam_Clamp_B750."));
                return "";
            }
        }

        #endregion
    }
}


