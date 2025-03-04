//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_Rod_Coupling_B655.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Rod_Coupling_B655

//   Author       :  Hema
//   Creation Date:  30.July.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.July.2012    Hema     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   26.Nov.2012     Hema     CR-CP-219113 Modified the code with SymbolGeomHelper
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
    [VariableOutputs]
    public class Bline_Rod_Coupling_B655 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Rod_Coupling"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "DC", "Diameter of Coupling", 0)]
        public InputDouble m_dDC;
        [InputDouble(3, "LC", "Length of Coupling", 0)]
        public InputDouble m_dLC;

        #endregion

        #region "Definitions of Aspects and their outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("ROD_COUPLING", "ROD_COUPLING")]

        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                
                double LC = m_dLC.Value;
                double DC = m_dDC.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                if (LC == 0.0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidLength, "Length value should not be equal to 0"));
                    return;
                }
                if (DC <= 0.0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDR, "Diameter value should not be less than or equal to 0"));
                    return;
                }
               

                //Ports
                Port inThdRH1 = new Port(OccurrenceConnection, part, "InThdRH1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = inThdRH1;

                Port inThdRH2 = new Port(OccurrenceConnection, part, "InThdRH2", new Position(0, 0, 0.001), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = inThdRH2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rodCoupling = (Projection3d)symbolGeometryHelper.CreateCylinder(null, (DC) / 2, LC);
                m_Symbolic.Outputs["ROD_COUPLING"] = rodCoupling;
            }
            catch  //General Unhandled exception 
            {
               
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_Rod_Coupling_B655."));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBlineFinish, "Bline Finish Code list value should be between 1 and 8"));
                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                bomDescription = catalogPart.PartDescription + ", Finish: " + finish;
                return bomDescription;
            }
            catch 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_Rod_Coupling_B655."));
                return "";
            }
        }
        #endregion
    }
}



