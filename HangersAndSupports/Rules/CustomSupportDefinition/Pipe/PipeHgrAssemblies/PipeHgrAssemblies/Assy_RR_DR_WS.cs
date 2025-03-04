//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_DR_WS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_DR_WS
//   Author       : BS
//   Creation Date: 07.Apr.2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.Apr.2013     BS      CR224472 Convert HS_Power_Assy to C# .Net 
//   22-02-2015      PVK     TR-CP-264951  Resolve coverity issues found in November 2014 report   
//   17-02-2015      Chethan TR-CP-266728  Assy_RR_DR_WS and Assy_RR_DR_LS not placing properly      
//   28-04-2015      PVK	 Resolve Coverity issues found in April
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    
    public class Assy_RR_DR_WS : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_DR_WS"
        //----------------------------------------------------------------------------------

        //Constants
        //For everything
        private const string HOR_SECTION = "hor_Section";
        private const string ROD1 = "rod1";
        private const string ROD2 = "rod2";
        private const string NUT1 = "nut1";
        private const string NUT2 = "nut2";
        private const string BOT_EYE_NUT1 = "bot_EYE_Nut1";
        private const string BOT_EYE_NUT2 = "bot_EYE_Nut2";
        private const string BOT_BEAM_ATT1 = "bot_Beam_Att1";
        private const string BOT_BEAM_ATT2 = "bot_Beam_Att2";

        //For LUG/CLEVIS
        private const string LUG = "lug";
        private const string CLEVIS = "clevis";
        private const string LUG2 = "lug2";
        private const string CLEVIS2 = "clevis2";
        private const string LUG_ROD3 = "lug_Rod3";
        private const string LUG_TB = "lug_tb";
        private const string LUG_NUT5 = "lug_Nut5";
        private const string LUG_NUT6 = "lug_Nut6";
        private const string LUG_ROD4 = "lug_Rod4";
        private const string LUG_TB2 = "lug_Tb2";
        private const string LUG_NUT7 = "lug_Nut7";
        private const string LUG_NUT8 = "lug_Nut8";

        //For BEAM_CLAMP
        private const string BEAM_CONNECTION = "beam_connection";
        private const string BEAM_CONNECTION2 = "beam_connection2";
        private const string BEAM_CLAMP = "beam_Clamp";
        private const string BEAM_CLAMP2 = "beam_Clamp2";
        private const string TB = "tb";
        private const string TB2 = "tb2";
        private const string ROD3 = "rod3";
        private const string ROD4 = "rod4";
        private const string NUT5 = "nut5";
        private const string NUT6 = "nut6";
        private const string NUT7 = "nut7";
        private const string NUT8 = "nut8";

        //For ROD_BEAM_ATT
        private const string BEAM_ATT = "beam_Att";
        private const string EYE_NUT = "eye_Nut";
        private const string BEAM_ATT2 = "beam_Att2";
        private const string EYE_NUT2 = "eye_Nut2";
        private const string ATT_ROD3 = "att_Rod3";
        private const string ATT_TB = "att_tb";
        private const string ATT_NUT5 = "att_Nut5";
        private const string ATT_NUT6 = "att_Nut6";
        private const string ATT_ROD4 = "att_Rod4";
        private const string ATT_TB2 = "att_Tb2";
        private const string ATT_NUT7 = "att_Nut7";
        private const string ATT_NUT8 = "att_Nut8";

        //For ROD_WASHER
        private const string WASHER = "washer";
        private const string WASH_NUT9 = "wash_Nut9";
        private const string WASH_NUT10 = "wash_Nut10";
        private const string WASHER2 = "washer2";
        private const string WASH_NUT11 = "wash_Nut11";
        private const string WASH_NUT12 = "wash_Nut12";
        private const string CONNECTION = "connection";
        private const string CONNECTION2 = "connection2";
        private const string WASH_ROD3 = "wash_Rod3";
        private const string WASH_TB = "wash_tb";
        private const string WASH_NUT5 = "wash_Nut5";
        private const string WASH_NUT6 = "wash_Nut6";
        private const string WASH_ROD4 = "wash_Rod4";
        private const string WASH_TB2 = "wash_tb2";
        private const string WASH_NUT7 = "wash_Nut7";
        private const string WASH_NUT8 = "wash_Nut8";

        //For ROD_NUT
        private const string NUT_CONNECTION = "nut_Connection";
        private const string NUT_NUT9 = "nut_Nut9";
        private const string NUT_NUT10 = "nut_Nut10";
        private const string NUT_CONNECTION2 = "nut_Connection2";
        private const string NUT_NUT11 = "nut_Nut11";
        private const string NUT_NUT12 = "nut_Nut12";
        private const string NUT_ROD3 = "nut_Rod3";
        private const string NUT_TB = "nut_tb";
        private const string NUT_NUT5 = "nut_Nut5";
        private const string NUT_NUT6 = "nut_Nut6";
        private const string NUT_ROD4 = "nut_Rod4";
        private const string NUT_TB2 = "nut_tb2";
        private const string NUT_NUT7 = "nut_Nut7";
        private const string NUT_NUT8 = "nut_Nut8";

        private string rodType = string.Empty,topType = string.Empty;
        private int turnbuckle;
        private double width1,width2,overhang,bottomRodLength;


        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    string sectionSize = ((PropertyValueString)support.GetPropertyValue("IJUAHgrWSize", "WSize")).PropValue;
                    rodType = ((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRodType", "ROD_TYPE")).PropValue;
                    topType = ((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyTopType", "TOP_TYPE")).PropValue;

                    turnbuckle = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTurnbuckle", "TURNBUCKLE")).PropValue;

                    overhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrTrapeze", "OVERHANG")).PropValue;
                    bottomRodLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyRR", "BOT_ROD_LENGTH")).PropValue;

                    string rodSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRR", "ROD_SIZE")).PropValue;
                    string beamClampSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRR", "BEAM_CLAMP_SIZE")).PropValue;

                    width1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W1")).PropValue;
                    width2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W2")).PropValue;

                    if (topType.ToUpper() == "ROD_CLEVIS_LUG")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));                            
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSize));
                            parts.Add(new PartInfo(CLEVIS, "Anvil_FIG299_" + rodSize));
                            parts.Add(new PartInfo(LUG2, "Anvil_FIG55S_" + rodSize));
                            parts.Add(new PartInfo(CLEVIS2, "Anvil_FIG299_" + rodSize));
                            parts.Add(new PartInfo(LUG_ROD3, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(LUG_TB, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(LUG_NUT5, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(LUG_NUT6, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(LUG_ROD4, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(LUG_TB2, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(LUG_NUT7, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(LUG_NUT8, "Anvil_HEX_NUT_" + rodSize));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSize));
                            parts.Add(new PartInfo(CLEVIS, "Anvil_FIG299_" + rodSize));
                            parts.Add(new PartInfo(LUG2, "Anvil_FIG55S_" + rodSize));
                            parts.Add(new PartInfo(CLEVIS2, "Anvil_FIG299_" + rodSize));

                        }
                    }

                    if (topType == "ROD_BEAM_CLAMP")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(BEAM_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(BEAM_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                            parts.Add(new PartInfo(BEAM_CLAMP2, "Anvil_FIG292_" + beamClampSize));
                            parts.Add(new PartInfo(ROD3, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(TB, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(NUT5, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT6, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(ROD4, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(TB2, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(NUT7, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT8, "Anvil_HEX_NUT_" + rodSize));
                        }
                        else
                        {
                            parts.Add(new PartInfo(BEAM_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(BEAM_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                            parts.Add(new PartInfo(BEAM_CLAMP2, "Anvil_FIG292_" + beamClampSize));
                        }
                    }
                    if (topType == "ROD_BEAM_ATT")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(EYE_NUT, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(ATT_ROD3, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(ATT_TB, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(ATT_NUT5, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(ATT_NUT6, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(ATT_ROD4, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(ATT_TB2, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(ATT_NUT7, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(ATT_NUT8, "Anvil_HEX_NUT_" + rodSize));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(EYE_NUT, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(EYE_NUT2, "Anvil_FIG290_" + rodSize));
                        }
                    }
                    if (topType == "ROD_WASHER")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(WASHER, "Anvil_FIG60_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT9, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT10, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASHER2, "Anvil_FIG60_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT11, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT12, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(WASH_ROD3, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(WASH_TB, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT5, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT6, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASH_ROD4, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(WASH_TB2, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT7, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT8, "Anvil_HEX_NUT_" + rodSize));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(WASHER, "Anvil_FIG60_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT9, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT10, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASHER2, "Anvil_FIG60_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT11, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(WASH_NUT12, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(CONNECTION2, "Log_Conn_Part_1"));
                        }
                    }
                    if (topType == "ROD_NUT")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT9, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT10, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT11, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT12, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_ROD3, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(NUT_TB, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT5, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT6, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_ROD4, "Anvil_FIG253_" + rodSize));
                            parts.Add(new PartInfo(NUT_TB2, "Anvil_FIG230_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT7, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT8, "Anvil_HEX_NUT_" + rodSize));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType + rodSize));
                            parts.Add(new PartInfo(ROD2, rodType + rodSize));
                            parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                            parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT9, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT10, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT11, "Anvil_HEX_NUT_" + rodSize));
                            parts.Add(new PartInfo(NUT_NUT12, "Anvil_HEX_NUT_" + rodSize));
                        }
                    }
                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 2;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //==========================
                //1. Load standard bounding box definition
                //==========================
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                string boundingBoxName;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBoxName = "BBSR";
                else
                    boundingBoxName = "BBR";

                //====== ======
                //2. retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry

                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | dHeight
                // |____________________|
                //        dWidth

                Double width, height;
                BoundingBox boundingBox = BoundingBoxHelper.GetBoundingBox(boundingBoxName);
                width = boundingBox.Width;
                height = boundingBox.Height;

                CommonAssembly commonAssembly = new CommonAssembly();
                
                string rightStructPort = string.Empty;
                string leftStructPort = string.Empty;

                double offset1 = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Y);

                if (offset1 < 0)
                {
                    leftStructPort = "Structure";
                    rightStructPort = "Struct_2";
                }
                else
                {
                    leftStructPort = "Struct_2";
                    rightStructPort = "Structure";
                }

                //====== ======
                //Set Values of Part Occurance Attributes
                //====== ======                
                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[HOR_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[HOR_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;

                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[HOR_SECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[HOR_SECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                //====== ======
                //Create Joints
                //====== ======               

                double lengthHor1 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", rightStructPort, PortDistanceType.Horizontal);
                double lengthHor2 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", leftStructPort, PortDistanceType.Horizontal);
                double lengthVert1 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", rightStructPort, PortDistanceType.Vertical);
                double lengthVert2 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", leftStructPort, PortDistanceType.Vertical);
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "BBRV_Low", PortAxisType.X, OrientationAlong.Direct);

                if ((width1 + width2) < width)
                {
                    width1 = (width + overhang) / 2;
                    width2 = (width + overhang) / 2;
                }
                SupportHelper.Support.SetPropertyValue(width1, "IJOAHgrAssy_RR_DR", "W1");
                SupportHelper.Support.SetPropertyValue(width2, "IJOAHgrAssy_RR_DR", "W2");

                double calc1 = Math.Cos(routeStructAngle) * width1;
                double calc2 = Math.Sin(routeStructAngle) * width1;
                double calc3 = Math.Cos(routeStructAngle) * width2;
                double calc4 = Math.Sin(routeStructAngle) * width2;

                double flangeThickness = 0;
                BusinessObject horizontalSectionPart = componentDictionary[HOR_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;

                string connectiontype = string.Empty;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    connectiontype = "Slab";
                else
                {
                    if ((SupportHelper.SupportingObjects.Count != 0))
                    {
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                            connectiontype = "Steel";
                        else
                            connectiontype = "Slab";
                    }
                    else
                        connectiontype = "Slab";
                }

                BusinessObject rod1 = componentDictionary[ROD1].GetRelationship("madeFrom", "part").TargetObjects[0];
                double rodDiameter = (double)((PropertyValueDouble)rod1.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue;
                BusinessObject nut1 = componentDictionary[NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                double nutThichness = (double)((PropertyValueDouble)nut1.GetPropertyValue("IJUAHgrAnvil_hex_nut", "T")).PropValue;


                double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                //figure out the orientation of the structure port
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);


                string configuration = string.Empty;
                double beamClampOffset;
                double beamLeftClampByStruct;
                double beamRightClampByStruct;
                double washerOffset;
                double nutOffset;

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI,7) / 2)    //The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2)
                    {
                        configuration = "1";
                        beamClampOffset = -lengthHor1 - width1;
                        washerOffset = lengthHor1 + width1;
                        nutOffset = lengthHor1 + width1;
                    }
                    else
                    {
                        configuration = "2";
                        beamClampOffset = lengthHor1 + width1;
                        washerOffset = -lengthHor1 - width1;
                        nutOffset = lengthHor1 + width1;
                    }
                }
                else    //The structure is oriented in the opposite direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2)
                    {
                        configuration = "3";
                        beamClampOffset = lengthHor1 + width1;
                        washerOffset = -lengthHor1 - width1;
                        nutOffset = lengthHor1 + width1;
                    }
                    else
                    {
                        configuration = "4";
                        beamClampOffset = -lengthHor1 - width1;
                        washerOffset = lengthHor1 + width1;
                        nutOffset = lengthHor1 + width1;
                    }
                }

                double byPointAngle3 = RefPortHelper.AngleBetweenPorts("BBRV_Low", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Direct);
                double distLeftClampRoute = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_Low", PortDistanceType.Horizontal);
                double distRightClampRoute = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_High", PortDistanceType.Horizontal);
                if (SupportHelper.SupportedObjects.Count == 1)
                {
                    if (Math.Abs(byPointAngle3) > Math.Round(Math.PI / 2, 7))    //The structure is oriented in the standard direction
                    {
                        beamLeftClampByStruct = width1;
                        beamRightClampByStruct = -width1;
                    }
                    else
                    {
                        beamLeftClampByStruct = -width1;
                        beamRightClampByStruct = width1;
                    }
                }
                else
                {
                    if ((Math.Abs(byPointAngle3) >= Math.Round(Math.PI / 2, 7)) && (Math.Abs(byPointAngle3) <= Math.Round(Math.PI / 2, 7)))//The structure is oriented in the standard direction
                    {
                        beamLeftClampByStruct = distLeftClampRoute + width2 - width1 + overhang / 2;
                        beamRightClampByStruct = -distRightClampRoute - width2 + width1 - overhang / 2;
                    }
                    else
                    {
                        beamLeftClampByStruct = -distLeftClampRoute - width2 + width1 - overhang / 2;
                        beamRightClampByStruct = distRightClampRoute + width2 - width1 + overhang / 2;
                    }
                }

                //Start the Joints here ********************************************************************************************************************
                //Add a Prismatic Joint defining the flexible bottom member
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    componentDictionary[HOR_SECTION].SetPropertyValue(width1 + width2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                else
                {
                    if (leftStructPort == rightStructPort)
                        componentDictionary[HOR_SECTION].SetPropertyValue(width1 + width2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                    else
                        componentDictionary[HOR_SECTION].SetPropertyValue(lengthHor1 + lengthHor2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                }

                //Create the Flexible (Prismatic) Joint between the ports of the top rods
                JointHelper.CreatePrismaticJoint(ROD1, "TopExThdRH", ROD1, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                JointHelper.CreatePrismaticJoint(ROD2, "TopExThdRH", ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                //Add a Vertical Joint to the Rods Z axes
                JointHelper.CreateGlobalAxesAlignedJoint(ROD1, "TopExThdRH", Axis.Z, Axis.Z);
                JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "TopExThdRH", Axis.Z, Axis.Z);

                if (topType.ToUpper() == "ROD_CLEVIS_LUG")
                {
                    //Add a revolute Joint between the lug hole and clevis pin
                    JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole", Axis.Y, Axis.X);
                    JointHelper.CreateRevoluteJoint(CLEVIS2, "Pin", LUG2, "Hole", Axis.Y, Axis.X);

                    //Add a rigid Joint between top of the rod and the Clevis
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", CLEVIS, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", CLEVIS2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + width1 + overhang, steelWidth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort)
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + width1, steelWidth / 2);
                        else
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelWidth / 2);
                    }

                    //Add Joints between the lug and the Structure
                    if (Configuration == 1)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - calc2);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -calc1 - width / 2);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 + calc1);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - calc2);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert2, 0, lengthHor2);
                        }
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, calc3 - width / 2);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 - calc3);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -lengthHor1);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc2, -calc1);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -calc1 - width / 2, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc2, -calc1);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert2, lengthHor2, 0);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc4, calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, calc3 - width / 2, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc4, calc3);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -lengthHor1, 0);
                        }
                    }
                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        bottomRodLength = bottomRodLength + height;                        
                        componentDictionary[LUG_ROD3].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");                        
                        componentDictionary[LUG_ROD4].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD3, "ExThdLH", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD4, "ExThdLH", Axis.Z, Axis.Z);

                        //Add a Rigid Joint between the bottom eye nut and the rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", LUG_ROD3, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", LUG_ROD4, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        //Add a Revolute Joint between the bottom eye nut and the beam attachment
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.XY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.XY, 0);
                        
                        //Add a Joint between bottom rod and Turnbuckle
                        JointHelper.CreateRigidJoint(LUG_ROD3, "ExThdLH", LUG_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD4, "ExThdLH", LUG_TB2, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Joint between top rod and Turnbuckle
                        JointHelper.CreateRigidJoint(LUG_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_TB2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(LUG_ROD4, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutThichness * 2, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD3, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutThichness * 2, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD3, "ExThdLH", LUG_NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", LUG_NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD4, "ExThdLH", LUG_NUT7, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", LUG_NUT8, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add a Rigid Joint between the bottom eye nut and the rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        //Add a Revolute Joint between the bottom eye nut and the beam attachment
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.XY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.XY, 0);

                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                    }
                }
                if (topType.ToUpper() == "ROD_BEAM_CLAMP")
                {
                    //We need to get the Flange Width.  However, it comes in metres, so we change it to inches.  Then we Int it to get rid of decimals.
                    //Then we add one so that the Flange Width of the clamp is bigger than the actual flange width, not smaller
                    double flangeWidth=0;
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeWidth = 3;
                    else
                    { 
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                                flangeWidth = Convert.ToInt32(SupportingHelper.SupportingObjectInfo(1).FlangeWidth * 1000 / 25.4) + 1;
                            else
                                flangeWidth = 3;
                        }
                        else
                            flangeWidth = 3;
                    }
                    //I know the lowest allowable FLANGE_W is 3 and the max is 15
                    if (flangeWidth < 3)
                        flangeWidth = 3;
                    else
                    {
                        if (flangeWidth > 15)
                            flangeWidth = 15;
                    }
                    CodelistItem flangeWidthCodelist = componentDictionary[BEAM_CLAMP].GetPropertyValue("IJOAHgrAnvil_FIG292", "FLANGE_WIDTH").PropertyInfo.CodeListInfo.GetCodelistItem(Convert.ToString(flangeWidth).Trim());

                    //We need to pass Flange Width into the Beam Clamp as a codelist Index so we get the Index from the Actual below.We do not need to pass in the Codelist Index.                    
                    componentDictionary[BEAM_CLAMP].SetPropertyValue(flangeWidthCodelist.Value, "IJOAHgrAnvil_FIG292", "FLANGE_WIDTH");
                    flangeWidthCodelist = componentDictionary[BEAM_CLAMP2].GetPropertyValue("IJOAHgrAnvil_FIG292", "FLANGE_WIDTH").PropertyInfo.CodeListInfo.GetCodelistItem(Convert.ToString(flangeWidth).Trim());
                    componentDictionary[BEAM_CLAMP2].SetPropertyValue(flangeWidthCodelist.Value, "IJOAHgrAnvil_FIG292", "FLANGE_WIDTH");

                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + width1 + overhang, steelWidth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort)
                            if (configuration == "1" || configuration == "3")
                                JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + width1 + overhang, steelWidth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + width1, steelWidth / 2);
                        else
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelWidth / 2);
                    }

                    //Connect Beam clamps to the structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, -beamLeftClampByStruct);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            if (configuration == "1" || configuration == "2")
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, beamClampOffset + width / 2, 0);
                            else
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, beamClampOffset - width / 2);
                        else //Two pieces of steel
                            JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                    }

                    //Add a rigid joint between the Beam clamps 2 to the structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, -beamRightClampByStruct);
                    else
                        JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", leftStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Planar Joint between top of the rod and the Beam Clamp
                    JointHelper.CreatePlanarJoint(ROD1, "TopExThdRH", BEAM_CLAMP, "InThdRH", Plane.XY, Plane.XY, 0);
                    JointHelper.CreatePlanarJoint(ROD2, "TopExThdRH", BEAM_CLAMP2, "InThdRH", Plane.XY, Plane.XY, 0);

                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                    //Add a Rigid Joint between the beam attachment and the route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc2, -calc1);
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc4, calc3);
                    }
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            if (configuration == "3" || configuration == "4")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc1, 0);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc3, 0);
                            }
                            else if (configuration == "1")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc1, 0);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc3, 0);
                            }
                            else
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc2, -calc1);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc4, calc3);
                            }
                        else //Two pieces of steel
                        {
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, lengthHor1, 0);
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -lengthHor2, 0);
                        }
                    }
                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[ROD3].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                        componentDictionary[ROD4].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");

                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(ROD3, "ExThdLH", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(ROD4, "ExThdLH", Axis.Z, Axis.Z);

                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(ROD3, "ExThdLH", TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "ExThdLH", TB2, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(TB2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add rigid joint between bottom rods and eye nuts
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD3, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD4, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(ROD3, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD3, "ExThdLH", NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "ExThdLH", NUT7, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", NUT8, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        // Add rigid joint between bottom rods and eye nuts
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                    }
                }

                if (topType.ToUpper() == "ROD_BEAM_ATT")
                {
                    //Add a revolute Joint between the beam attachment and eye nut
                    JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.Y, Axis.X);
                    JointHelper.CreateRevoluteJoint(EYE_NUT2, "Eye", BEAM_ATT2, "Pin", Axis.Y, Axis.X);

                    //Add a rigid Joint between top of the rod and the eye nut
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", EYE_NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + width1 + overhang, steelWidth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort)
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + width1, steelWidth / 2);
                        else
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelWidth / 2);
                    }

                    //Add Joints between the beam attachment and the Structure
                    if (Configuration == 1)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2.25 + calc4, calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2.25 + calc3, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2.25 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2.25 + calc4, calc3);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -lengthHor1, 0);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc2, -calc1);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc2, -calc1);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert2, lengthHor2, 0);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 - calc3);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 + calc3);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -lengthHor1);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - calc2);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -width / 2 - calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - calc2);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert2, 0, lengthHor2);
                        }
                    }

                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[ATT_ROD3].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                        componentDictionary[ATT_ROD4].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");

                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD3, "ExThdLH", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD4, "ExThdLH", Axis.Z, Axis.Z);

                        //Add a Rigid Joint between the bottom eye nut and the rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ATT_ROD3, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ATT_ROD4, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        //Add a Revolute Joint between the bottom eye nut and the beam attachment
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(BOT_BEAM_ATT1, "Structure", HOR_SECTION, "BeginCap", Plane.XY, Plane.ZX, 0);//
                        JointHelper.CreatePlanarJoint(BOT_BEAM_ATT2, "Structure", HOR_SECTION, "EndCap", Plane.XY, Plane.ZX, 0);//

                        //Add a Joint between bottom rod and Turnbuckle
                        JointHelper.CreateRigidJoint(ATT_ROD3, "ExThdLH", ATT_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "ExThdLH", ATT_TB2, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        //Add a Joint between top rod and Turnbuckle
                        JointHelper.CreateRigidJoint(ATT_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_TB2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(ATT_ROD3, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD3, "ExThdLH", ATT_NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", ATT_NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "ExThdLH", ATT_NUT7, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", ATT_NUT8, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(BOT_BEAM_ATT1, "Structure", HOR_SECTION, "BeginCap", Plane.XY, Plane.ZX, 0);
                        JointHelper.CreatePlanarJoint(BOT_BEAM_ATT2, "Structure", HOR_SECTION, "EndCap", Plane.XY, Plane.ZX, 0);

                        //Add a Revolute Joint between the bottom eye nut and the beam attachment
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                        //Add a Rigid Joint between the bottom eye nut and the rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                    }
                }

                if (topType.ToUpper() == "ROD_WASHER")
                {

                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeThickness = 0.02;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                                flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                            else
                                flangeThickness = 0.02;
                        }
                        else
                            flangeThickness = 0.02;
                    }

                    //Add a Planar Joint between top of the rod and the Beam Clamp
                    JointHelper.CreatePlanarJoint(ROD1, "TopExThdRH", CONNECTION, "Connection", Plane.XY, Plane.NegativeXY, 0);
                    JointHelper.CreatePlanarJoint(ROD2, "TopExThdRH", CONNECTION2, "Connection", Plane.XY, Plane.NegativeXY, 0);

                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + width1 + overhang, steelWidth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + width1, steelWidth / 2);
                        else
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelWidth / 2);
                    }

                    //Add a Rigid Joint between the connection and the Route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                        JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                        JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                    }

                    //Add a Prismatic Joint between the lug and the Structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(WASHER, "Structure", ROD2, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            JointHelper.CreateRigidJoint(WASHER, "Structure", ROD2, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                        else //Two pieces of steel
                            JointHelper.CreateRigidJoint(WASHER, "Structure", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0.01, 0, 0);
                    }

                    //Add a Prismatic Joint between the lug and the Structure
                    JointHelper.CreateRigidJoint(WASHER2, "Structure", ROD1, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);

                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                    //Add a Rigid Joint between the beam attachment and the route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc2, -calc1);
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc4, calc3);
                    }
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                        {
                            if (configuration == "3" || configuration == "4")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc1, 0);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc3, 0);
                            }
                            else if (configuration == "1")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc1, 0);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc3, 0);
                            }
                            else
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc2, -calc1);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc4, calc3);
                            }
                        }
                        else //Two pieces of steel
                        {
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, lengthHor1, 0);
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -lengthHor2, 0);
                        }
                    }

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", WASH_NUT9, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", WASH_NUT10, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", WASH_NUT11, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", WASH_NUT12, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);

                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[WASH_ROD3].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                        componentDictionary[WASH_ROD4].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD3, "ExThdLH", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD4, "ExThdLH", Axis.Z, Axis.Z);

                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(WASH_ROD3, "ExThdLH", WASH_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "ExThdLH", WASH_TB2, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(WASH_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_TB2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between bottom eye nut and wash rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", WASH_ROD3, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", WASH_ROD4, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(WASH_ROD3, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD3, "ExThdLH", WASH_NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", WASH_NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "ExThdLH", WASH_NUT7, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", WASH_NUT8, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                    }
                }

                if (topType.ToUpper() == "ROD_NUT")
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeThickness = 0.02;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                                flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                            else
                                flangeThickness = 0.02;
                        }
                        else
                            flangeThickness = 0.02;
                    }

                    //Add a Planar Joint between top of the rod and the Beam Clamp
                    JointHelper.CreatePlanarJoint(ROD1, "TopExThdRH", NUT_CONNECTION, "Connection", Plane.XY, Plane.NegativeXY, 0);
                    JointHelper.CreatePlanarJoint(ROD2, "TopExThdRH", NUT_CONNECTION2, "Connection", Plane.XY, Plane.NegativeXY, 0);

                    //Add a Rigid Joint between the connection and the Route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                    }
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                        {
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, nutOffset, 0);
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);
                        }
                    }

                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + width1 + overhang, steelWidth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + width1, steelWidth / 2);
                        else
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelWidth / 2);
                    }

                    //Add a Rigid Joint between the beam attachment and the route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc2, -calc1);
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc4, calc3);
                    }
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                        {
                            if (configuration == "1") 
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc1, 0);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc3, 0);
                            }
                            else if (configuration == "3" || configuration == "4")
                            {
                                if (HgrCompareDoubleService.cmpdbl(byPointAngle3 , 0)==true)
                                {
                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc1, 0);
                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc3, 0);                                    
                                }
                                else
                                {
                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc1, 0);
                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc3, 0);                                    
                                }
                            }
                            else
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 + calc2, -calc1);
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, width / 2 - calc4, calc3);
                            }
                        }
                        else //Two pieces of steel
                        {
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, lengthHor1, 0);
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -lengthHor2, 0);
                        }
                    }

                    //Add a Revolute Joint between the bottom eye nut and the beam attachment
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                    //Add a joint between the top nuts and the top rods
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT9, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT10, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT_NUT11, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT_NUT12, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);

                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[NUT_ROD3].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                        componentDictionary[NUT_ROD4].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD3, "ExThdLH", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD4, "ExThdLH", Axis.Z, Axis.Z);

                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(NUT_ROD3, "ExThdLH", NUT_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "ExThdLH", NUT_TB2, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(NUT_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_TB2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between bottom eye nut and wash rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", NUT_ROD3, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", NUT_ROD4, "ExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(NUT_ROD3, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD3, "ExThdLH", NUT_NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT_NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "ExThdLH", NUT_NUT7, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", NUT_NUT8, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add rigid joint between bottom rods and eye nuts
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThichness, 0, 0);
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    for (int iIndex = 1; iIndex <= SupportHelper.SupportedObjects.Count; iIndex++) // partindex, routeindex
                    {
                        routeConnections.Add(new ConnectionInfo(HOR_SECTION, iIndex));
                    }

                    //Return the collection of Route connection information
                    return routeConnections;
                }
                catch (Exception e)
                {

                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Structure connection information
                    int supportingObjects;
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        supportingObjects = 1;
                    else
                        supportingObjects = SupportHelper.SupportingObjects.Count;

                    if (topType == "ROD_CLEVIS_LUG")
                    {
                        structConnections.Add(new ConnectionInfo(LUG, 1)); // partindex, routeindex
                        if (supportingObjects > 1)
                            structConnections.Add(new ConnectionInfo(LUG2, 2));
                        else
                            structConnections.Add(new ConnectionInfo(LUG2, 1));
                    }
                    else if (topType == "ROD_BEAM_CLAMP")
                    {
                        structConnections.Add(new ConnectionInfo(BEAM_CLAMP, 1));
                        if (supportingObjects > 1)
                            structConnections.Add(new ConnectionInfo(BEAM_CLAMP2, 2));
                        else
                            structConnections.Add(new ConnectionInfo(BEAM_CLAMP2, 1));
                    }
                    else if (topType == "ROD_BEAM_ATT")
                    {
                        structConnections.Add(new ConnectionInfo(BEAM_ATT, 1));
                        if (supportingObjects > 1)
                            structConnections.Add(new ConnectionInfo(BEAM_ATT2, 2));
                        else
                            structConnections.Add(new ConnectionInfo(BEAM_ATT2, 1));
                    }
                    else if (topType == "ROD_WASHER")
                    {
                        structConnections.Add(new ConnectionInfo(WASHER, 1));
                        if (supportingObjects > 1)
                            structConnections.Add(new ConnectionInfo(WASHER2, 2));
                        else
                            structConnections.Add(new ConnectionInfo(WASHER2, 1));
                    }
                    else if (topType == "ROD_NUT")
                    {
                        structConnections.Add(new ConnectionInfo(NUT_NUT9, 1));
                        if (supportingObjects > 1)
                            structConnections.Add(new ConnectionInfo(NUT_NUT11, 2));
                        else
                            structConnections.Add(new ConnectionInfo(NUT_NUT11, 1));
                        
                    }

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }

            }
        }
        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {
                if (eMirrorPlane == MirrorPlane.XZPlane)
                {
                    //Interchange the values of W1 and W2
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    double w1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W1")).PropValue;
                    double w2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W2")).PropValue;
                    support.SetPropertyValue(w2, "IJOAHgrAssy_RR_DR", "W1");
                    support.SetPropertyValue(w1, "IJOAHgrAssy_RR_DR", "W2");

                }                
                return CurrentMirrorToggleValue;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Mirrored Configuration." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
            
