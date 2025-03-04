//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_DR_WS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_DR_WS
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01-10-2015      Vinay   DI-CP-276996	Update HS_Assembly2 to use RichSteel
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   21/03/2016      Vinay   TR-CP-288920	Issues found in HS_Assembly_V2
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

                    width1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W1")).PropValue;
                    width2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W2")).PropValue;

                    if (topType.ToUpper() == "ROD_CLEVIS_LUG")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));                            
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType ));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut" ));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(LUG, "Anv10_ShortStructLug"));
                            parts.Add(new PartInfo(CLEVIS, "Anv10_ClevisWithPin"));
                            parts.Add(new PartInfo(LUG2, "Anv10_ShortStructLug"));
                            parts.Add(new PartInfo(CLEVIS2, "Anv10_ClevisWithPin"));
                            parts.Add(new PartInfo(LUG_ROD3, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(LUG_TB, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(LUG_NUT5, "Anv10_HexNut"));
                            parts.Add(new PartInfo(LUG_NUT6, "Anv10_HexNut"));
                            parts.Add(new PartInfo(LUG_ROD4, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(LUG_TB2, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(LUG_NUT7, "Anv10_HexNut"));
                            parts.Add(new PartInfo(LUG_NUT8, "Anv10_HexNut"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(LUG, "Anv10_ShortStructLug"));
                            parts.Add(new PartInfo(CLEVIS, "Anv10_ClevisWithPin"));
                            parts.Add(new PartInfo(LUG2, "Anv10_ShortStructLug"));
                            parts.Add(new PartInfo(CLEVIS2, "Anv10_ClevisWithPin"));

                        }
                    }

                    if (topType == "ROD_BEAM_CLAMP")
                    {
                        double offset1 = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.Y);
                        int refIndex1 = 1;
                        int refIndex2 = 1;

                        if (SupportHelper.SupportingObjects.Count > 1)
                        {
                            if (offset1 < 0)
                            {
                                refIndex1 = 2;
                                refIndex2 = 1;
                            }
                            else
                            {
                                refIndex1 = 1;
                                refIndex2 = 2;
                            }
                        }

                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(BEAM_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(BEAM_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BEAM_CLAMP, "Anv10_MBeamClamp_292", "", refIndex1));
                            parts.Add(new PartInfo(BEAM_CLAMP2, "Anv10_MBeamClamp_292", "", refIndex2));
                            parts.Add(new PartInfo(ROD3, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(TB, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(NUT5, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT6, "Anv10_HexNut"));
                            parts.Add(new PartInfo(ROD4, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(TB2, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(NUT7, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT8, "Anv10_HexNut"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(BEAM_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(BEAM_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BEAM_CLAMP, "Anv10_MBeamClamp_292", "", refIndex1));
                            parts.Add(new PartInfo(BEAM_CLAMP2, "Anv10_MBeamClamp_292", "", refIndex2));
                        }
                    }
                    if (topType == "ROD_BEAM_ATT")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BEAM_ATT, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(EYE_NUT, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(ATT_ROD3, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(ATT_TB, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(ATT_NUT5, "Anv10_HexNut"));
                            parts.Add(new PartInfo(ATT_NUT6, "Anv10_HexNut"));
                            parts.Add(new PartInfo(ATT_ROD4, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(ATT_TB2, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(ATT_NUT7, "Anv10_HexNut"));
                            parts.Add(new PartInfo(ATT_NUT8, "Anv10_HexNut"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BEAM_ATT, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(EYE_NUT, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(EYE_NUT2, "Anv10_EyeNut"));
                        }
                    }
                    if (topType == "ROD_WASHER")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(WASHER, "Anv10_WasherPlate"));
                            parts.Add(new PartInfo(WASH_NUT9, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASH_NUT10, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASHER2, "Anv10_WasherPlate"));
                            parts.Add(new PartInfo(WASH_NUT11, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASH_NUT12, "Anv10_HexNut"));
                            parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(WASH_ROD3, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(WASH_TB, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(WASH_NUT5, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASH_NUT6, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASH_ROD4, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(WASH_TB2, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(WASH_NUT7, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASH_NUT8, "Anv10_HexNut"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(WASHER, "Anv10_WasherPlate"));
                            parts.Add(new PartInfo(WASH_NUT9, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASH_NUT10, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASHER2, "Anv10_WasherPlate"));
                            parts.Add(new PartInfo(WASH_NUT11, "Anv10_HexNut"));
                            parts.Add(new PartInfo(WASH_NUT12, "Anv10_HexNut"));
                            parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(CONNECTION2, "Log_Conn_Part_1"));
                        }
                    }
                    if (topType == "ROD_NUT")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT9, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_NUT10, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT11, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_NUT12, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_ROD3, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(NUT_TB, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(NUT_NUT5, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_NUT6, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_ROD4, "Anv10_RodETRL"));
                            parts.Add(new PartInfo(NUT_TB2, "Anv10_Turnbuckle"));
                            parts.Add(new PartInfo(NUT_NUT7, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_NUT8, "Anv10_HexNut"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT1, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_EYE_NUT2, "Anv10_EyeNut"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anv10_WBABolt"));
                            parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT9, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_NUT10, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_CONNECTION2, "Log_Conn_Part_1"));
                            parts.Add(new PartInfo(NUT_NUT11, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT_NUT12, "Anv10_HexNut"));
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

                if (SupportHelper.SupportingObjects.Count > 1)
                {
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
                }
                else
                {
                    leftStructPort = "Structure";
                    rightStructPort = "Structure";
                }

                //====== ======
                //Set Values of Part Occurance Attributes
                ////====== ======                

                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

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
                double supportingSectionDepth = 0;
                BusinessObject horizontalSectionPart = componentDictionary[HOR_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;

                string connectiontype = string.Empty;
                
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                        connectiontype = "Steel";
                    else
                        connectiontype = "Slab";
                }
                else
                    connectiontype = "Slab";

                BusinessObject rod1 = componentDictionary[ROD1].GetRelationship("madeFrom", "part").TargetObjects[0];
                double rodDiameter = (double)((PropertyValueDouble)rod1.GetPropertyValue("IJUAhsRodDiameter","RodDiameter")).PropValue;
                BusinessObject nut1 = componentDictionary[NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                double nutThichness = (double)((PropertyValueDouble)nut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;


                double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                //figure out the orientation of the structure port
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);


                string configuration = string.Empty;
                double beamClampOffset;
                double beamLeftClampByStruct;
                double beamRightClampByStruct;

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI/2,7) / 2)    //The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI/2, 7) / 2)
                    {
                        configuration = "1";
                        beamClampOffset = -lengthHor1 - width1;
                    }
                    else
                    {
                        configuration = "2";
                        beamClampOffset = lengthHor1 + width1;
                    }
                }
                else    //The structure is oriented in the opposite direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI/2, 7) / 2)
                    {
                        configuration = "3";
                        beamClampOffset = lengthHor1 + width1;
                    }
                    else
                    {
                        configuration = "4";
                        beamClampOffset = -lengthHor1 - width1;
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
                JointHelper.CreatePrismaticJoint(ROD1, "RodEnd1", ROD1, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                JointHelper.CreatePrismaticJoint(ROD2, "RodEnd1", ROD2, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                //Add a Vertical Joint to the Rods Z axes
                JointHelper.CreateGlobalAxesAlignedJoint(ROD1, "RodEnd1", Axis.Z, Axis.Z);
                JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "RodEnd1", Axis.Z, Axis.Z);

                if (topType.ToUpper() == "ROD_CLEVIS_LUG")
                {
                    //Add a revolute Joint between the lug hole and clevis pin
                    JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole1", Axis.Y, Axis.Y);
                    JointHelper.CreateRevoluteJoint(CLEVIS2, "Pin", LUG2, "Hole1", Axis.Y, Axis.Y);

                    //Add a rigid Joint between top of the rod and the Clevis
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CLEVIS, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", CLEVIS2, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

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
                    if (Configuration == 2)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - calc2);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -calc1 - width / 2);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 + calc1);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - calc2);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert2, 0, lengthHor2);
                        }
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, calc3 - width / 2);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 - calc3);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -lengthHor1);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc2, -calc1);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -calc1 - width / 2, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc2, -calc1);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert2, lengthHor2, 0);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc4, calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, calc3 - width / 2, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc4, calc3);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -lengthHor1, 0);
                        }
                    }
                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        bottomRodLength = bottomRodLength + height;
                        componentDictionary[LUG_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[LUG_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD4, "RodEnd1", Axis.Z, Axis.Z);

                        //Add a Rigid Joint between the bottom eye nut and the rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", LUG_ROD3, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", LUG_ROD4, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        //Add a Revolute Joint between the bottom eye nut and the beam attachment
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                        
                        //Add a Joint between bottom rod and Turnbuckle
                        JointHelper.CreateRigidJoint(LUG_ROD3, "RodEnd2", LUG_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD4, "RodEnd2", LUG_TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Joint between top rod and Turnbuckle
                        if (Configuration == 1 || Configuration == 3)
                        {
                            JointHelper.CreateRigidJoint(LUG_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);
                            JointHelper.CreateRigidJoint(LUG_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);
                        }
                        else if (Configuration == 2 || Configuration == 4)
                        {
                            JointHelper.CreateRigidJoint(LUG_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            JointHelper.CreateRigidJoint(LUG_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }

                        //Add a Rigid Joint between the nuts and the rods
                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        JointHelper.CreateRigidJoint(LUG_ROD4, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutThichness, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[LUG_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition3 = (opening3 - rodTakeOut3) / 2 + shapeLength3;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[LUG_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition4 = (opening4 - rodTakeOut4) / 2 + shapeLength4;
                        JointHelper.CreateRigidJoint(LUG_ROD3, "RodEnd2", LUG_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LUG_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3+ nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD4, "RodEnd2", LUG_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", LUG_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                    }
                    else
                    {
                        //Add a Rigid Joint between the bottom eye nut and the rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Revolute Joint between the bottom eye nut and the beam attachment
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.X);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.X);

                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1  + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;

                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutThichness, 0, 0);
                    }
                }
                if (topType.ToUpper() == "ROD_BEAM_CLAMP")
                {
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
                        JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -beamLeftClampByStruct);
                    else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", "BBRV_Low", Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 - calc3 , 0);
                    }
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            if (configuration == "1" || configuration == "2")
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, beamClampOffset + width / 2);
                            else
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, beamClampOffset - width / 2);
                        else //Two pieces of steel
                            JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }

                    //Add a rigid joint between the Beam clamps 2 to the structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -beamRightClampByStruct);
                    else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", "BBRV_Low", Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 + calc1, 0);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            if (configuration == "1" || configuration == "2")
                                JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -beamClampOffset + width / 2);
                            else
                                JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, beamClampOffset + width / 2);
                        else //Two pieces of steel
                            JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    }

                    //Add a Planar Joint between top of the rod and the Beam Clamp
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", BEAM_CLAMP, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", BEAM_CLAMP2, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");

                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(ROD4, "RodEnd1", Axis.Z, Axis.Z);

                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(ROD3, "RodEnd2", TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "RodEnd2", TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add rigid joint between bottom rods and eye nuts
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ROD3, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ROD4, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a revolute Joint between the bottom nut and the beam attachments
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);

                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        JointHelper.CreateRigidJoint(ROD4, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutThichness, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition3 = (opening3 - rodTakeOut3) / 2 + shapeLength3;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition4 = (opening4 - rodTakeOut4) / 2 + shapeLength4;
                        JointHelper.CreateRigidJoint(ROD3, "RodEnd2", NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "RodEnd2", NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                    }
                    else
                    {
                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);

                        //Add a revolute Joint between the bottom nut and the beam attachments
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.X);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.X);

                        // Add rigid joint between bottom rods and eye nuts
                        if (Configuration == 1)
                        {
                            JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);
                            JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }
                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                    }
                }

                if (topType.ToUpper() == "ROD_BEAM_ATT")
                {
                    //Add a revolute Joint between the beam attachment and eye nut
                    JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.Y, Axis.Y);
                    JointHelper.CreateRevoluteJoint(EYE_NUT2, "Eye", BEAM_ATT2, "Pin", Axis.Y, Axis.Y);

                    //Add a rigid Joint between top of the rod and the eye nut
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", EYE_NUT2, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

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
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 - calc4, calc3);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 - calc3, 0);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 - calc3, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 + calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 - calc4, 0);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, lengthHor1, 0);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 + calc2, -calc1);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 + calc1, 0);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 + calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 - calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1, width / 2 + calc2, 0);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert2, -lengthHor2 , 0);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, calc3, -width / 2 + calc4);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 - calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 - calc3);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 + calc3);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, -calc3, -width / 2 + calc4);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, 0, -lengthHor1);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, -calc1, -width / 2 - calc2);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            if (SupportHelper.SupportedObjects.Count > 1)
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, 0, width / 2 - calc1 + overhang);
                            else
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, 0,  calc1 - overhang);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, -width / 2 - calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, calc1, -width / 2 - calc2);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert2, 0, lengthHor2);
                        }
                    }

                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[ATT_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[ATT_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");

                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD4, "RodEnd1", Axis.Z, Axis.Z);

                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(ATT_ROD3, "RodEnd2", ATT_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "RodEnd2", ATT_TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between top rods and Turnbuckle
                        if (Configuration == 1 || Configuration == 3)
                        {
                            JointHelper.CreateRigidJoint(ATT_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);
                            JointHelper.CreateRigidJoint(ATT_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);
                        }
                        else if (Configuration == 2 || Configuration == 4)
                        {
                            JointHelper.CreateRigidJoint(ATT_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            JointHelper.CreateRigidJoint(ATT_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }

                        // Add rigid joint between bottom rods and eye nuts
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ATT_ROD3, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ATT_ROD4, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a revolute Joint between the bottom nut and the beam attachments
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.X);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.X);

                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        JointHelper.CreateRigidJoint(ATT_ROD4, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutThichness, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[ATT_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition3 = (opening3 - rodTakeOut3) / 2 + shapeLength3;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[ATT_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition4 = (opening4 - rodTakeOut4) / 2 + shapeLength4;
                        JointHelper.CreateRigidJoint(ATT_ROD3, "RodEnd2", ATT_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", ATT_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "RodEnd2", ATT_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", ATT_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                    }
                    else
                    {
                        //Add planar joints for beam attachment to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);

                        //Add a revolute Joint between the bottom nut and the beam attachments
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.X);
                        JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.X);

                        // Add rigid joint between bottom rods and eye nuts
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                    }
                }

                if (topType.ToUpper() == "ROD_WASHER")
                {

                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        flangeThickness = 0.02;
                        supportingSectionDepth = 0.1016;
                    }
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                            {
                                flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                                supportingSectionDepth = SupportingHelper.SupportingObjectInfo(1).Depth;
                            }
                            else
                            {
                                flangeThickness = 0;
                                supportingSectionDepth = 0;
                            }
                        }
                        else
                        {
                            flangeThickness = 0;
                            supportingSectionDepth = 0;
                        }
                    }
                    //Add a Rigid Joint between Connection and the RodEnd
                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", ROD1, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", ROD2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

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
                    if (Configuration == 1)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - calc4, calc3);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - calc3 , 0);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + calc3, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, -width / 2 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + calc4, 0);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, lengthHor1, 0);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + calc2, -calc1);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + calc1, 0);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - calc2, 0);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, lengthVert2 + 0.05 + flangeThickness / 2, -lengthHor2, 0);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, calc3, -width / 2 + calc4);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 - calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc3);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc3);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc4);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -lengthHor1);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, -calc1, -width / 2 - calc2);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc1);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 - calc1);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 - +0.05 + flangeThickness / 2, 0, -width / 2 - calc1);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 - calc2);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert2 + 0.05 + flangeThickness / 2, 0, lengthHor2);
                        }
                    }

                    //Add a Rigid Joint between the lug and the Structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(WASHER, "Port2", ROD2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 - 0.05, 0, 0);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            JointHelper.CreateRigidJoint(WASHER, "Port2", ROD2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 - 0.05, 0, 0);
                        else //Two pieces of steel
                            JointHelper.CreateRigidJoint(WASHER, "Port2", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0.01, 0, 0);
                    }

                    //Add a Prismatic Joint between the lug and the Structure
                    JointHelper.CreateRigidJoint(WASHER2, "Port2", ROD1, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 - 0.05, 0, 0);

                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.X);
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.X);

                    //Add planar joints for beam attachment to horizontal section
                    JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                    JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);
                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", WASH_NUT9, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.03 + nutThichness, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", WASH_NUT10, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.015 + nutThichness, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", WASH_NUT11, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.03 + nutThichness, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", WASH_NUT12, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.015 + nutThichness, 0, 0);

                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[WASH_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[WASH_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD4, "RodEnd1", Axis.Z, Axis.Z);

                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(WASH_ROD3, "RodEnd2", WASH_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "RodEnd2", WASH_TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(WASH_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between bottom eye nut and wash rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", WASH_ROD3, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", WASH_ROD4, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        JointHelper.CreateRigidJoint(WASH_ROD4, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutThichness, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[WASH_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition3 = (opening3 - rodTakeOut3) / 2 + shapeLength3;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[WASH_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition4 = (opening4 - rodTakeOut4) / 2 + shapeLength4;

                        JointHelper.CreateRigidJoint(WASH_ROD3, "RodEnd2", WASH_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", WASH_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "RodEnd2", WASH_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", WASH_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                    }
                    else
                    {
                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a Rigid Joint between the bottom nut and the rod
                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
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
                                flangeThickness = 0;
                        }
                        else
                            flangeThickness = 0;
                    }
                    //Add a Rigid Joint between Connection and ROdEnd
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT_CONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", NUT_CONNECTION2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Calculate the position of Top Nut
                    double topNutPosition = lengthVert1 + 0.05 + flangeThickness / 2; //Handled for placement on W beam 
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference)
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if (SupportingHelper.SupportingObjectInfo(1).FaceNumber == 514)
                                topNutPosition = lengthVert1 + 0.05 - flangeThickness / 2;
                            else if (SupportingHelper.SupportingObjectInfo(1).FaceNumber == 513)
                                topNutPosition = lengthVert1 + 0.05 + flangeThickness / 2;
                        }
                    }

                    //Add a Rigid Joint between the connection and the Route
                    if (Configuration == 1)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 - calc4, calc3);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 - calc3, 0);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 + calc3, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, -width / 2 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 + calc4, 0);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, lengthHor1, 0);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 + calc2, -calc1);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 + calc1, 0);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, width / 2 - calc2, 0);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, topNutPosition, -lengthHor2, 0);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, calc3, -width / 2 + calc4);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 - calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 + calc3);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 + calc3);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 + calc4);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -lengthHor1);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, -calc1, -width / 2 - calc2);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 + calc1);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 - calc1);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, -width / 2 - calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 - calc2);
                            else //Two pieces of steel
                                JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, 0, lengthHor2);
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

                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.X);
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.X);

                    //Add planar joints for beam attachment to horizontal section
                    JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.NegativeZX, Plane.XY, 0);
                    JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.NegativeXY, 0);

                    //Add a joint between the top nuts and the top rods
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT_NUT9, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05 + nutThichness, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT_NUT10, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 + 0.05 , 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", NUT_NUT11, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05 + nutThichness, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", NUT_NUT12, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 + 0.05 , 0, 0);

                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[NUT_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[NUT_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD4, "RodEnd1", Axis.Z, Axis.Z);

                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(NUT_ROD3, "RodEnd2", NUT_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "RodEnd2", NUT_TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(NUT_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between bottom eye nut and wash rod
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", NUT_ROD3, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", NUT_ROD4, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        JointHelper.CreateRigidJoint(NUT_ROD4, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutThichness, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[NUT_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength3 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition3 = (opening3 - rodTakeOut3) / 2 + shapeLength3;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[NUT_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength4 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition4 = (opening4 - rodTakeOut4) / 2 + shapeLength4;

                        JointHelper.CreateRigidJoint(NUT_ROD3, "RodEnd2", NUT_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "RodEnd2", NUT_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition4 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition3 + nutThichness, 0, 0);
                    }
                    else
                    {
                        //Add rigid joint between bottom rods and eye nuts
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "RodEnd", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "RodEnd", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the bottom nut and the rod

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut1 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = overLength1 + shapeLength1;

                        //Getting the nut positions for EyeNut1
                        BusinessObject eyeNut2 = componentDictionary[BOT_EYE_NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double overLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)eyeNut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = overLength1 + shapeLength1;
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
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
            
