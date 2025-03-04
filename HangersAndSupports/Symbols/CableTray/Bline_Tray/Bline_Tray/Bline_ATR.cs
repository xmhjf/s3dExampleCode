//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_ATR.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_ATR
//   Author       :  Vijaya
//   Creation Date:  25.July.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   25.July.2012    Vijaya     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   23.Nov.2012     Hema       CR-CP-219113 Modified the code with SymbolGeomHelper
//   10.Jan.2010     Hema       CR-CP-219113 Changed the implementation of WeightCG
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;


namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.Support.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------


    public class Bline_ATR : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        private const double ATRconstant1 = 0.015875;
        private const double ATRconstant2 = 0.6096;
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_ATR"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "DR", "Diameter of the Rod", 0)]
        public InputDouble m_dDiameter;
        [InputDouble(3, "Length", "Length of the Rod", 0)]
        public InputDouble m_dLength;


        #endregion

        #region "Definitions of Aspects and their outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("ROD", "ROD")]

        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                double length = m_dLength.Value;
                double DR = m_dDiameter.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================                
                if (length == 0.0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidLength, "Length value should not be equal to 0"));
                    return;
                }
                if (DR <= 0.0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDR, "Diameter value should not be less than or equal to 0"));
                    return;
                }

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "BotExThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "TopExThdRH", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rod = (Projection3d)symbolGeometryHelper.CreateCylinder(null, DR / 2, length);
                m_Symbolic.Outputs["ROD"] = rod;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_ATR."));
                    return;
                }
            }
        }
        #endregion

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;

                PropertyValueCodelist finishCodeList = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAHgrBlineFinish", "Finish");
                long finishValue = finishCodeList.PropValue;

                if (finishValue < 0 || finishValue > 8)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBlineFinish, "B-line Finish Code list value should be between 1 and 8"));

                if (finishValue == 0)
                    finishValue = 1;

                string finish = finishCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)finishValue).DisplayName;
                string size = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUABlineATR", "SIZE")).PropValue;
                double DValue = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUABlineATR", "Dia")).PropValue;
                double D_USER = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOABlineATR", "D_User")).PropValue;
                double DR = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUABlineATR", "DR")).PropValue;
                if (D_USER > 0.0 && lengthValue < D_USER * 2.0)
                    DValue = lengthValue / 2.0;
                else
                {
                    if (D_USER > 0.0)
                        DValue = D_USER;
                    else
                        if ((DR <= ATRconstant1) && (lengthValue <= ATRconstant2))
                            DValue = lengthValue / 2.0;
                }
                string D = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, DValue, UnitName.DISTANCE_INCH);
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_INCH);
                bomString = catalogPart.PartDescription + ", Size: " + size + "," + " Cut Length: " + length + ", Finish: " + finish;
                return bomString;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_ATR."));
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
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                double weightPerUnitLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrBlineWeightPerL", "Weight_Per_Length")).PropValue;

                weight = weightPerUnitLength * length;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of Bline_ATR."));

            }
        }
        #endregion
    }
}




