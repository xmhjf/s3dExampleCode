﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_SR_CL.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_SR_CL
//   Author       : Rajeswari
//   Creation Date: 17-April-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 17-April-2013  Rajeswari CR-CP-224484 C#.Net HS_Assembly Project Creation  
//   22-02-2015       PVK   TR-CP-264951  Resolve coverity issues found in November 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class Assy_RR_SR_CL : CustomSupportDefinition
    {
        // For everything
        private const string PIPE_CLAMP = "PIPE_CLAMP";
        private const string ROD1 = "ROD1";
        private const string NUT1 = "NUT1";
        private const string NUT2 = "NUT2";

        // For C_CLAMP
        private const string C_CLAMP = "C_CLAMP";
        private const string ROD2 = "ROD2";
        private const string TB = "TB";
        private const string NUT3 = "NUT3";
        private const string NUT4 = "NUT4";

        // For LUG/CLEVIS
        private const string LUG = "LUG";
        private const string CLEVIS = "CLEVIS";
        private const string LUG_ROD2 = "LUG_ROD2";
        private const string LUG_TB = "LUG_TB";
        private const string LUG_NUT3 = "LUG_NUT3";
        private const string LUG_NUT4 = "LUG_NUT4";

        // For BEAM_CLAMP
        private const string BEAM_CLAMP = "BEAM_CLAMP";

        // For ROD_BEAM_ATT
        private const string BEAM_ATT = "BEAM_ATT";
        private const string EYE_NUT = "EYE_NUT";
        private const string ATT_ROD2 = "ATT_ROD2";
        private const string ATT_TB = "ATT_TB";
        private const string ATT_NUT3 = "ATT_NUT3";
        private const string ATT_NUT4 = "ATT_NUT4";

        // For ROD_WASHER
        private const string WASHER = "WASHER";
        private const string NUT5 = "NUT5";
        private const string NUT6 = "NUT6";
        private const string CONNECTION = "CONNECTION";
        private const string WASH_ROD2 = "WASH_ROD2";
        private const string WASH_TB = "WASH_TB";
        private const string WASH_NUT3 = "WASH_NUT3";
        private const string WASH_NUT4 = "WASH_NUT4";

        // For ROD_NUT
        private const string NUT_CONNECTION = "NUT_CONNECTION";
        private const string NUT_NUT5 = "NUT_NUT5";
        private const string NUT_NUT6 = "NUT_NUT6";
        private const string NUT_ROD2 = "NUT_ROD2";
        private const string NUT_TB = "NUT_TB";
        private const string NUT_NUT3 = "NUT_NUT3";
        private const string NUT_NUT4 = "NUT_NUT4";

        private const double CONST_INCH = 1000.0 / 25.4;
        String topType, rodType, beamClampSize, rodSize;
        int turnbuckle;
        Double botRodLength;

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

                    rodType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRodType", "ROD_TYPE")).PropValue;
                    topType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyTopType", "TOP_TYPE")).PropValue;
                    turnbuckle = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTurnbuckle", "TURNBUCKLE")).PropValue;
                    botRodLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyRR", "BOT_ROD_LENGTH")).PropValue;
                    rodSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRR", "ROD_SIZE")).PropValue;
                    beamClampSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRR", "BEAM_CLAMP_SIZE")).PropValue;
                    
                    PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    double pipeDiameter = routeInfo.OutsideDiameter;

                    string temp = rodType.Substring(9, 3);
                    double rodDiameter=0.0;
                    string rodSizeValue = string.Empty;

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass anvil_FIG260PartClass = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG260");
                    ReadOnlyCollection<BusinessObject> anvil_FIG260Parts = anvil_FIG260PartClass.Parts;

                    foreach (BusinessObject part in anvil_FIG260Parts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue > pipeDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue < pipeDiameter + 0.001)
                        {
                            rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrAnvil_FIG260", "A")).PropValue;
                            break;
                        }
                    }

                    if (temp == "146")
                    {
                        PartClass anvil_FIG146PartClass = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG146");
                        anvil_FIG260Parts = anvil_FIG146PartClass.Parts;
                    }
                    else
                    {
                        PartClass anvil_FIG140PartClass = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG140");
                        anvil_FIG260Parts = anvil_FIG140PartClass.Parts;
                    }

                    foreach (BusinessObject part in anvil_FIG260Parts)
                     {
                         if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue > rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue < rodDiameter + 0.001)
                         {
                             rodSizeValue = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrSize", "SIZE")).PropValue;
                             break;
                         }
                     }

                    if (rodSizeValue == string.Empty)
                    {
                        // Set it up for C-Clamp Usage ****************************************************************************************************
                        if (topType == "ROD_C_CLAMP")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(C_CLAMP, "Anvil_FIG86_" + rodSize));
                                parts.Add(new PartInfo(ROD2, "Anvil_FIG253_" + rodSize));
                                parts.Add(new PartInfo(TB, "Anvil_FIG230_" + rodSize));
                                parts.Add(new PartInfo(NUT3, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT4, "Anvil_HEX_NUT_" + rodSize));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(C_CLAMP, "Anvil_FIG86_" + rodSize));
                            }
                        }

                        // Set it up for Lug/Clevis Usage ****************************************************************************************************
                        if (topType == "ROD_CLEVIS_LUG")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSize));
                                parts.Add(new PartInfo(CLEVIS, "Anvil_FIG299_" + rodSize));
                                parts.Add(new PartInfo(LUG_ROD2, "Anvil_FIG253_" + rodSize));
                                parts.Add(new PartInfo(LUG_TB, "Anvil_FIG230_" + rodSize));
                                parts.Add(new PartInfo(LUG_NUT3, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(LUG_NUT4, "Anvil_HEX_NUT_" + rodSize));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSize));
                                parts.Add(new PartInfo(CLEVIS, "Anvil_FIG299_" + rodSize));
                            }
                        }

                        // Set it up for Rod/Beam Clamp Usage ****************************************************************************************************
                        if (topType == "ROD_BEAM_CLAMP")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                                parts.Add(new PartInfo(ROD2, "Anvil_FIG253_" + rodSize));
                                parts.Add(new PartInfo(TB, "Anvil_FIG230_" + rodSize));
                                parts.Add(new PartInfo(NUT3, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT4, "Anvil_HEX_NUT_" + rodSize));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                            }
                        }

                        // Set it up for ROD_BEAM_ATT Usage ****************************************************************************************************
                        if (topType == "ROD_BEAM_ATT")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSize));
                                parts.Add(new PartInfo(EYE_NUT, "Anvil_FIG290_" + rodSize));
                                parts.Add(new PartInfo(ATT_ROD2, "Anvil_FIG253_" + rodSize));
                                parts.Add(new PartInfo(ATT_TB, "Anvil_FIG230_" + rodSize));
                                parts.Add(new PartInfo(ATT_NUT3, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(ATT_NUT4, "Anvil_HEX_NUT_" + rodSize));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSize));
                                parts.Add(new PartInfo(EYE_NUT, "Anvil_FIG290_" + rodSize));
                            }
                        }

                        // Set it up for ROD_WASHER Usage ****************************************************************************************************
                        if (topType == "ROD_WASHER")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(WASHER, "Anvil_FIG60_" + rodSize));
                                parts.Add(new PartInfo(NUT5, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT6, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(WASH_ROD2, "Anvil_FIG253_" + rodSize));
                                parts.Add(new PartInfo(WASH_TB, "Anvil_FIG230_" + rodSize));
                                parts.Add(new PartInfo(WASH_NUT3, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(WASH_NUT4, "Anvil_HEX_NUT_" + rodSize));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(WASHER, "Anvil_FIG60_" + rodSize));
                                parts.Add(new PartInfo(NUT5, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT6, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                            }
                        }

                        // Set it up for ROD_NUT Usage ****************************************************************************************************
                        if (topType == "ROD_NUT")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(NUT_NUT5, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT_NUT6, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT_ROD2, "Anvil_FIG253_" + rodSize));
                                parts.Add(new PartInfo(NUT_TB, "Anvil_FIG230_" + rodSize));
                                parts.Add(new PartInfo(NUT_NUT3, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT_NUT4, "Anvil_HEX_NUT_" + rodSize));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSize));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(NUT_NUT5, "Anvil_HEX_NUT_" + rodSize));
                                parts.Add(new PartInfo(NUT_NUT6, "Anvil_HEX_NUT_" + rodSize));
                            }
                        }
                    }
                    else
                    {
                        if (topType == "ROD_C_CLAMP")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(C_CLAMP, "Anvil_FIG86_" + rodSizeValue));
                                parts.Add(new PartInfo(ROD2, "Anvil_FIG253_" + rodSizeValue));
                                parts.Add(new PartInfo(TB, "Anvil_FIG230_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT3, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT4, "Anvil_HEX_NUT_" + rodSizeValue));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(C_CLAMP, "Anvil_FIG86_" + rodSizeValue));
                            }
                        }

                        // Set it up for Lug/Clevis Usage ****************************************************************************************************
                        if (topType == "ROD_CLEVIS_LUG")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSizeValue));
                                parts.Add(new PartInfo(CLEVIS, "Anvil_FIG299_" + rodSizeValue));
                                parts.Add(new PartInfo(LUG_ROD2, "Anvil_FIG253_" + rodSizeValue));
                                parts.Add(new PartInfo(LUG_TB, "Anvil_FIG230_" + rodSizeValue));
                                parts.Add(new PartInfo(LUG_NUT3, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(LUG_NUT4, "Anvil_HEX_NUT_" + rodSizeValue));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSizeValue));
                                parts.Add(new PartInfo(CLEVIS, "Anvil_FIG299_" + rodSizeValue));
                            }
                        }

                        // Set it up for Rod/Beam Clamp Usage ****************************************************************************************************
                        if (topType == "ROD_BEAM_CLAMP")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                                parts.Add(new PartInfo(ROD2, "Anvil_FIG253_" + rodSizeValue));
                                parts.Add(new PartInfo(TB, "Anvil_FIG230_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT3, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT4, "Anvil_HEX_NUT_" + rodSizeValue));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                            }
                        }

                        // Set it up for ROD_BEAM_ATT Usage ****************************************************************************************************
                        if (topType == "ROD_BEAM_ATT")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSizeValue));
                                parts.Add(new PartInfo(EYE_NUT, "Anvil_FIG290_" + rodSizeValue));
                                parts.Add(new PartInfo(ATT_ROD2, "Anvil_FIG253_" + rodSizeValue));
                                parts.Add(new PartInfo(ATT_TB, "Anvil_FIG230_" + rodSizeValue));
                                parts.Add(new PartInfo(ATT_NUT3, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(ATT_NUT4, "Anvil_HEX_NUT_" + rodSizeValue));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSizeValue));
                                parts.Add(new PartInfo(EYE_NUT, "Anvil_FIG290_" + rodSizeValue));
                            }
                        }

                        // Set it up for ROD_WASHER Usage ****************************************************************************************************
                        if (topType == "ROD_WASHER")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(WASHER, "Anvil_FIG60_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT5, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT6, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(WASH_ROD2, "Anvil_FIG253_" + rodSizeValue));
                                parts.Add(new PartInfo(WASH_TB, "Anvil_FIG230_" + rodSizeValue));
                                parts.Add(new PartInfo(WASH_NUT3, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(WASH_NUT4, "Anvil_HEX_NUT_" + rodSizeValue));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(WASHER, "Anvil_FIG60_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT5, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT6, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                            }
                        }

                        // Set it up for ROD_NUT Usage ****************************************************************************************************
                        if (topType == "ROD_NUT")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(NUT_NUT5, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_NUT6, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_ROD2, "Anvil_FIG253_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_TB, "Anvil_FIG230_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_NUT3, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_NUT4, "Anvil_HEX_NUT_" + rodSizeValue));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG260"));
                                parts.Add(new PartInfo(ROD1, rodType + rodSizeValue));
                                parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(NUT_NUT5, "Anvil_HEX_NUT_" + rodSizeValue));
                                parts.Add(new PartInfo(NUT_NUT6, "Anvil_HEX_NUT_" + rodSizeValue));
                            }
                        }
                    }
                  
                    // Return the collection of Catalog Parts
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject pipeClampPart = componentDictionary[PIPE_CLAMP].GetRelationship("madeFrom", "part").TargetObjects[0];

                // ====== ======
                // Create Joints
                // ====== ======
                // Add the intelligence to determine which side of the pipe the steel is when placing by point
                Double byPointAngle1, byPointAngle2, nutPosition1, flangeThickness=0.0;
                Plane planeA, planeB, planeC, planeD;
                Axis axisA, axisB;

                byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI, 7) / 2.0)  // The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2.0)
                    {
                        planeA = Plane.XY;
                        planeB = Plane.ZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.ZX;
                    }
                    else
                    {
                        planeA = Plane.XY;
                        planeB = Plane.NegativeZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.NegativeZX;
                    }
                }
                else   // The structure is oriented in the opposite direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2.0)
                    {
                        planeA = Plane.XY;
                        planeB = Plane.ZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.NegativeZX;
                    }
                    else
                    {
                        planeA = Plane.XY;
                        planeB = Plane.NegativeZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.ZX;
                    }
                }

                nutPosition1 = ((double)((PropertyValueDouble)pipeClampPart.GetPropertyValue("IJUAHgrAnvil_FIG260", "B")).PropValue - (double)((PropertyValueDouble)pipeClampPart.GetPropertyValue("IJUAHgrTake_Out", "TAKE_OUT")).PropValue);

                // Start the Joints here ********************************************************************************************************************
                // Add a Vertical Joint to the Rod Z axis

                JointHelper.CreateGlobalAxesAlignedJoint(ROD1, "TopExThdRH", Axis.Z, Axis.Z);

                // Create the Flexible (Prismatic) Joint between the ports of the bottom rod

                JointHelper.CreatePrismaticJoint(ROD1, "TopExThdRH", ROD1, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                // Set it up for C-Clamp Usage
                if (topType == "ROD_C_CLAMP")
                {
                    // Raise the error for Place By reference
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "C-Clamp Top Type can not be placed by 'Place By Reference' Command", "", "Assy_RR_SR_CL.cs", 583);
                        return;
                    }
                    if ((SupportHelper.SupportingObjects.Count != 0))
                    {
                        flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                    }

                    componentDictionary[C_CLAMP].SetPropertyValue(flangeThickness, "IJOAHgrAnvil_FIG86", "FLANGE_T");

                    // Joint from Pipe Clamp to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        JointHelper.CreatePrismaticJoint(C_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0);

                        JointHelper.CreateRevoluteJoint(ROD1, "TopExThdRH", C_CLAMP, "InThdRH",Axis.Z,Axis.Z);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", planeA, planeB, axisA, axisB, 0, 0, 0);

                        JointHelper.CreatePlanarJoint(C_CLAMP, "Structure", "-1", "Structure", planeC, planeD, 0);

                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", C_CLAMP, "InThdRH", Plane.ZX, Plane.NegativeYZ, Axis.X, Axis.Y, 0, 0, 0);
                    }
                   
                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[ROD2].SetPropertyValue(botRodLength, "IJUAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "ExThdLH", Axis.Z, Axis.Z);

                        // Add a Rigid Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRigidJoint(ROD2, "ExThdRH", PIPE_CLAMP, "Hole", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(ROD2, "ExThdLH", TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD2, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD2, "ExThdLH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);

                        JointHelper.CreateRigidJoint(ROD2, "ExThdLH", NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        // Add a revolute Joint between bottom rod and Low Eye Nut

                        JointHelper.CreateRevoluteJoint(ROD1, "BotExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);
                    }
                }

                // Set it up for Clevis/Lug Usage
                if (topType == "ROD_CLEVIS_LUG")
                {
                    // Joint from Pipe Clamp to Pipe

                    JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Planar Joint between the lug and the Structure

                    JointHelper.CreatePlanarJoint(LUG, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    // Add a revolute Joint between the lug hole and clevis pin

                    JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole", Axis.X, Axis.Y);

                    // Add a Joint between top of the rod and the Clevis

                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", CLEVIS, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", CLEVIS, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[LUG_ROD2].SetPropertyValue(botRodLength, "IJUAHgrOccLength", "Length");

                        // add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD2, "ExThdLH", 0, 0);

                        // Add a revolute Joint between bottom rod and Low Eye Nut

                        JointHelper.CreateRevoluteJoint(LUG_ROD2, "ExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(LUG_ROD2, "ExThdLH", LUG_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(LUG_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(LUG_ROD2, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(LUG_ROD2, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);

                        JointHelper.CreateRigidJoint(LUG_ROD2, "ExThdLH", LUG_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", LUG_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        // Add a revolute Joint between bottom rod and Low Eye Nut

                        JointHelper.CreateRevoluteJoint(ROD1, "BotExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);
                    }
                }

                // Set it up for Beam Clamp Usage
                if (topType == "ROD_BEAM_CLAMP")
                {
                    Double flangeWidth;

                    // We need to get the Flange Width.  However, it comes in metres, so we change it to inches.  Then we Int it to get rid of decimals.
                    // Then we add one so that the Flange Width of the clamp is bigger than the actual flange width, not smaller
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeWidth = 3;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            flangeWidth = (int)(SupportingHelper.SupportingObjectInfo(1).Width * CONST_INCH) + 1.0;
                        }
                        else
                            flangeWidth = 3;
                    }

                    if (flangeWidth < 3)
                        flangeWidth = 3;
                    else
                    {
                        if (flangeWidth > 15)
                            flangeWidth = 15;
                    }

                    CodelistItem beamClampFWCodelist = componentDictionary[BEAM_CLAMP].GetPropertyValue("IJOAHgrAnvil_FIG292", "FLANGE_WIDTH").PropertyInfo.CodeListInfo.GetCodelistItem(flangeWidth.ToString().Trim());
                    componentDictionary[BEAM_CLAMP].SetPropertyValue(beamClampFWCodelist.Value, "IJOAHgrAnvil_FIG292", "FLANGE_WIDTH");

                    // Joint from Pipe Clamp to Pipe

                    // Changed Rigid to Prismatic for TR 87508
                    JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                    // Add a revolute Joint between top of the rod and the C Clamp
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0);
                    else
                        // Changed Rigid to Prismatic for TR 87508
                        JointHelper.CreatePrismaticJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.NegativeY, 0, 0);

                    // Add a revolute Joint between top of the rod and the Beam Clamp
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRevoluteJoint(ROD1, "TopExThdRH", BEAM_CLAMP, "InThdRH", Axis.Z, Axis.Z);
                    else
                        // Changed Rigid to Prismatic for TR 87508
                        JointHelper.CreateSphericalJoint(ROD1, "TopExThdRH", BEAM_CLAMP, "InThdRH");

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[ROD2].SetPropertyValue(botRodLength, "IJUAHgrOccLength", "Length");

                        // dd a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "ExThdLH", Axis.Z, Axis.Z);

                        // Add a revolute Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRevoluteJoint(ROD2, "ExThdRH", PIPE_CLAMP, "Hole",Axis.X, Axis.Y);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(ROD2, "ExThdLH", TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD2, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD2, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);

                        JointHelper.CreateRigidJoint(ROD2, "ExThdLH", NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add a revolute Joint between bottom rod and Clevis Hanger
                        JointHelper.CreateRevoluteJoint(ROD1, "BotExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);
                    }
                }

                // Set it up for Beam Att/Eye Nut Usage
                if (topType == "ROD_BEAM_ATT")
                {
                    // Joint from Pipe Clamp to Pipe

                    JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Planar Joint between the lug and the Structure

                    JointHelper.CreatePlanarJoint(BEAM_ATT, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    // Add a revolute Joint between the lug hole and clevis pin

                    JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.X, Axis.Y);

                    // Add a Joint between top of the rod and the Clevis
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                    if (turnbuckle == 1)
                    {
                        componentDictionary[ATT_ROD2].SetPropertyValue(botRodLength, "IJUAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD2, "ExThdLH", Axis.Z, Axis.Z);

                        // Add a revolute Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRevoluteJoint(ATT_ROD2, "ExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(ATT_ROD2, "ExThdLH", ATT_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(ATT_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ATT_ROD2, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ATT_ROD2, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);

                        JointHelper.CreateRigidJoint(ATT_ROD2, "ExThdLH", ATT_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", ATT_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add a revolute Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRevoluteJoint(ROD1, "BotExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);
                    }
                }

                // Set it up for Rod/Washer Usage
                if (topType == "ROD_WASHER")
                {
                    // Joint from Pipe Clamp to Pipe

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Prismatic Joint between the lug and the Structure

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                         JointHelper.CreatePrismaticJoint(WASHER, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, 0, 0);
                    else
                          JointHelper.CreatePlanarJoint(WASHER, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                     // Add a revolute Joint between top of the rod and the Connection

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRevoluteJoint(ROD1, "TopExThdRH", CONNECTION, "Connection", Axis.Z, Axis.Z);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", CONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Rigid Joint between the bottom nut and the rod

                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);

                    // Add a Rigid Joint between the middle nut and the rod

                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);

                    // Create the rigid joint to locate the connection objecct

                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", WASHER, "OppStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.05, 0, 0);

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[WASH_ROD2].SetPropertyValue(botRodLength, "IJUAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD2, "ExThdLH", Axis.Z, Axis.Z);

                        // Add a revolute Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRevoluteJoint(WASH_ROD2, "ExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(WASH_ROD2, "ExThdLH", WASH_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(WASH_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(WASH_ROD2, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(WASH_ROD2, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);

                        JointHelper.CreateRigidJoint(WASH_ROD2, "ExThdLH", WASH_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", WASH_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add a revolute Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRevoluteJoint(ROD1, "BotExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);
                    }
                }

                // Set it up for Rod/NUT Usage
                if (topType == "ROD_NUT")
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        // NO STRUCTURE IS SELCTED, SHOULD GO WITH DEFAULT SIZE
                        flangeThickness = 0.0;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                        }
                        else
                            flangeThickness = 0.0;
                    }
                    // Joint from Pipe Clamp to Pipe

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Prismatic Joint between the lug and the Structure

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(NUT_CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, -flangeThickness / 2.0 - 0.05, 0);
                    else
                        JointHelper.CreatePlanarJoint(NUT_CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, -flangeThickness / 2.0 - 0.05);

                    // Add a revolute Joint between top of the rod and the C Clamp

                    JointHelper.CreateRevoluteJoint(ROD1, "TopExThdRH", NUT_CONNECTION, "Connection", Axis.Z, Axis.Z);

                    // Add a Rigid Joint between the bottom nut and the rod

                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);

                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[NUT_ROD2].SetPropertyValue(botRodLength, "IJUAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD2, "ExThdLH", Axis.Z, Axis.Z);

                        // Add a revolute Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRevoluteJoint(NUT_ROD2, "ExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(NUT_ROD2, "ExThdLH", NUT_TB, "InThdLH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(NUT_TB, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(NUT_ROD2, "ExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(NUT_ROD2, "ExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);

                        JointHelper.CreateRigidJoint(NUT_ROD2, "ExThdLH", NUT_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    }
                    else
                    {
                        //Add a revolute Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRevoluteJoint(ROD1, "BotExThdRH", PIPE_CLAMP, "Hole", Axis.X, Axis.Y);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1 / 2, 0, 0);

                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -nutPosition1, 0, 0);
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
                    routeConnections.Add(new ConnectionInfo(PIPE_CLAMP, 1)); // partindex, routeindex

                    //Return the collection of Route connection information.
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
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    if (topType == "ROD_C_CLAMP")
                        structConnections.Add(new ConnectionInfo(C_CLAMP, 1)); // partindex, routeindex
                    if (topType == "ROD_CLEVIS_LUG")
                        structConnections.Add(new ConnectionInfo(LUG, 1));
                    if (topType == "ROD_BEAM_CLAMP")
                        structConnections.Add(new ConnectionInfo(BEAM_CLAMP, 1));
                    if (topType == "ROD_BEAM_ATT")
                        structConnections.Add(new ConnectionInfo(BEAM_ATT, 1));
                    if (topType == "ROD_WASHER")
                        structConnections.Add(new ConnectionInfo(WASHER, 1));
                    if (topType == "ROD_NUT")
                        structConnections.Add(new ConnectionInfo(NUT_CONNECTION, 1));

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

      }
}
