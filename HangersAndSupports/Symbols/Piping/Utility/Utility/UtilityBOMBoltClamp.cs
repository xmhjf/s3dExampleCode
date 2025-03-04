//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   UtilityBOMBoltClamp.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.UtilityBOMBoltClamp
//   Author       :  Rajeswari
//   Creation Date:  31/10/2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31/10/2012   Rajeswari   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

#region "ICustomHgrBOMDescription Members"
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

    [SymbolVersion("1.0.0.0")]
    public class UtilityBOMBoltClamp : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_GEN_3_BOLT_CLAMP"))
                {
                    double RValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "R")).PropValue;
                    double AValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "A")).PropValue;
                    double KValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "K")).PropValue;
                    double HValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "H")).PropValue;
                    double boltDiameterValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "BOLT_DIA")).PropValue;
                    double boltLValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "BOLT_L")).PropValue;
                    double TValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "T")).PropValue;
                    double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "W")).PropValue;

                    string R = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, RValue, UnitName.DISTANCE_INCH);
                    string A = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, AValue, UnitName.DISTANCE_INCH);
                    string K = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, KValue, UnitName.DISTANCE_INCH);
                    string H = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, HValue, UnitName.DISTANCE_INCH);
                    string boltDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, boltDiameterValue, UnitName.DISTANCE_INCH);
                    string boltL = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, boltLValue, UnitName.DISTANCE_INCH);
                    string T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_INCH);
                    string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);

                    PropertyValueCodelist bomOPTCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_3_BOLT_CLAMP", "OPT_BOM");

                    string optBom = bomOPTCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(bomOPTCodelist.PropValue).DisplayName;

                    bomString = "Custom Clamp, R=" + R + ", A=" + A + ", K=" + K +
                            ", H=" + H + ", Bolt=" + boltDiameter + "x" +
                            boltL + ", T=" + T + ", W=" + W;

                    if (optBom == "No")
                    {
                        bomString = "";
                    }
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_GEN_4_BOLT_CLAMP"))
                {
                    double RValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "R")).PropValue;
                    double AValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "A")).PropValue;
                    double KValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "K")).PropValue;
                    double HValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "H")).PropValue;
                    double boltDiameterValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "BOLT_DIA")).PropValue;
                    double boltLValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "BOLT_L")).PropValue;
                    double TValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "T")).PropValue;
                    double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "W")).PropValue;

                    string R = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, RValue, UnitName.DISTANCE_INCH);
                    string A = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, AValue, UnitName.DISTANCE_INCH);
                    string K = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, KValue, UnitName.DISTANCE_INCH);
                    string H = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, HValue, UnitName.DISTANCE_INCH);
                    string boltDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, boltDiameterValue, UnitName.DISTANCE_INCH);
                    string boltL = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, boltLValue, UnitName.DISTANCE_INCH);
                    string T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_INCH);
                    string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);

                    PropertyValueCodelist bomOPTCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_4_BOLT_CLAMP", "OPT_BOM");

                    string optBom = bomOPTCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(bomOPTCodelist.PropValue).DisplayName;

                    bomString = "Custom Clamp, R=" + R + ", A=" + A + ", K=" + K +
                            ", H=" + H + ", Bolt=" + boltDiameter + "x" +
                            boltL + ", T=" + T + ", W=" + W;

                    if (optBom == "No")
                    {
                        bomString = "";
                    }
                }
                else if (oSupportOrComponent.SupportsInterface("IJOAHgrUtility_4_LIN_BOLT_CL"))
                {
                    double RValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "R")).PropValue;
                    double AValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "A")).PropValue;
                    double KValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "K")).PropValue;
                    double HValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "H")).PropValue;
                    double boltDiameterValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "BOLT_DIA")).PropValue;
                    double boltLValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "BOLT_L")).PropValue;
                    double TValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "T")).PropValue;
                    double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "W")).PropValue;

                    string R = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, RValue, UnitName.DISTANCE_INCH);
                    string A = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, AValue, UnitName.DISTANCE_INCH);
                    string K = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, KValue, UnitName.DISTANCE_INCH);
                    string H = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, HValue, UnitName.DISTANCE_INCH);
                    string boltDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, boltDiameterValue, UnitName.DISTANCE_INCH);
                    string boltL = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, boltLValue, UnitName.DISTANCE_INCH);
                    string T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_INCH);
                    string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);

                    PropertyValueCodelist bomOPTCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_4_LIN_BOLT_CL", "OPT_BOM");

                    string optBom = bomOPTCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(bomOPTCodelist.PropValue).DisplayName;

                    bomString = "Custom Clamp, R=" + R + ", A=" + A + ", K=" + K +
                            ", H=" + H + ", Bolt=" + boltDiameter + "x" +
                            boltL + ", T=" + T + ", W=" + W;

                    if (optBom == "No")
                    {
                        bomString = "";
                    }
                }
                return bomString;
            }
            catch 
            {
                ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of UtilityBOMBoltClamp"));
                return "";
            }
        }
    }
}
#endregion
