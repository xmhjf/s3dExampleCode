//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_SS4_V.cs
//   Power_Assy,Ingr.SP3D.Content.Support.Rules.Assy_SS4_V
//   Author       :  Vijaya
//   Creation Date:  20.Mar.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20.Mar.2013    Vijaya   CR-CP-224472-Initial Creation
//   22/02/2015     PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

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
    public class Assy_SS4_V : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string ROD_CLEVIS_LUG = "ROD_CLEVIS_LUG"; // 1
        private const string ROD_BEAM_CLAMP = "ROD_BEAM_CLAMP"; //2
        private const string ROD_BEAM_ATT = "ROD_BEAM_ATT"; // 3
        private const string ROD_WASHER = "ROD_WASHER"; // 4
        private const string ROD_NUT = " ROD_NUT "; // 5
        private string[] Shoepartkeys = new string[15];

        // For everything
        private const string HOR_SECTION = "HOR_SECTION "; // 1
        private const string ROD1 = "ROD1"; // 2
        private const string ROD2 = "ROD2"; //3
        private const string NUT1 = "NUT1"; // 4
        private const string NUT2 = " NUT2 "; // 5
        private const string BOT_EYE_NUT1 = " BOT_EYE_NUT1 "; // 6
        private const string BOT_EYE_NUT2 = " BOT_EYE_NUT2 "; // 7
        private const string BOT_BEAM_ATT1 = " BOT_BEAM_ATT1 "; // 6
        private const string BOT_BEAM_ATT2 = " BOT_BEAM_ATT2 "; // 7

        //For LUG/CLEVIS
        private const string LUG = "LUG "; // 8
        private const string CLEVIS = "CLEVIS "; // 9
        private const string LUG2 = "LUG2 "; // 10
        private const string CLEVIS2 = "CLEVIS2 "; // 11
        private const string LUG_ROD3 = "LUG_ROD3 "; // 12
        private const string LUG_TB = "LUG_TB "; // 13
        private const string LUG_NUT5 = " LUG_NUT5 "; // 14
        private const string LUG_NUT6 = " LUG_NUT6 "; // 15
        private const string LUG_ROD4 = " LUG_ROD4 "; // 16
        private const string LUG_TB2 = " LUG_TB2 "; // 17
        private const string LUG_NUT7 = " LUG_NUT7 "; // 18
        private const string LUG_NUT8 = " LUG_NUT8 "; // 19

        //For BEAM_CLAMP
        private const string BEAM_CLAMP = " BEAM_CLAMP "; // 8
        private const string BEAM_CLAMP2 = " BEAM_CLAMP2 "; // 9
        private const string TB = " TB "; // 10
        private const string TB2 = " TB2 "; // 11
        private const string ROD3 = " ROD3 "; // 12
        private const string ROD4 = " ROD4 "; // 13
        private const string NUT5 = " NUT5 "; // 14
        private const string NUT6 = " NUT6 "; // 15
        private const string NUT7 = " NUT7 "; // 16
        private const string NUT8 = " NUT8 "; // 17

        //For ROD_BEAM_ATT
        private const string BEAM_ATT = " BEAM_ATT "; // 8
        private const string EYE_NUT = " EYE_NUT "; // 9
        private const string BEAM_ATT2 = " BEAM_ATT2 "; // 10
        private const string EYE_NUT2 = " EYE_NUT2 "; // 11
        private const string ATT_ROD3 = " ATT_ROD3 "; // 12
        private const string ATT_TB = " ATT_TB "; // 13
        private const string ATT_NUT5 = " ATT_NUT5 "; // 14
        private const string ATT_NUT6 = " ATT_NUT6 "; // 15
        private const string ATT_ROD4 = " ATT_ROD4 "; // 16
        private const string ATT_TB2 = " ATT_TB2 "; // 17
        private const string ATT_NUT7 = " ATT_NUT7 "; // 18
        private const string ATT_NUT8 = " ATT_NUT8 "; // 19

        //For ROD_WASHER
        private const string WASHER = " WASHER "; // 8
        private const string WASH_NUT9 = " WASH_NUT9 "; // 9
        private const string WASH_NUT10 = " WASH_NUT10 "; // 10
        private const string WASHER2 = " WASHER2 "; // 11
        private const string WASH_NUT11 = " WASH_NUT11 "; // 12
        private const string WASH_NUT12 = " WASH_NUT12 "; // 13
        private const string CONNECTION = " CONNECTION "; // 14
        private const string CONNECTION2 = " CONNECTION2 "; // 15
        private const string WASH_ROD3 = " WASH_ROD3 "; // 16
        private const string WASH_TB = " WASH_TB "; // 17
        private const string WASH_NUT5 = " WASH_NUT5 "; // 18
        private const string WASH_NUT6 = " WASH_NUT6 "; // 19
        private const string WASH_ROD4 = " WASH_ROD4 "; // 20
        private const string WASH_TB2 = " WASH_TB2 "; // 21
        private const string WASH_NUT7 = " WASH_NUT7 "; // 22
        private const string WASH_NUT8 = " WASH_NUT8 "; // 23

        //For ROD_NUT
        private const string NUT_CONNECTION = " NUT_CONNECTION "; // 8
        private const string NUT_NUT9 = " NUT_NUT9 "; // 9
        private const string NUT_NUT10 = " NUT_NUT10 "; // 10
        private const string NUT_CONNECTION2 = " NUT_CONNECTION2 "; // 11
        private const string NUT_NUT11 = " NUT_NUT11 "; // 12
        private const string NUT_NUT12 = " NUT_NUT12 "; // 13
        private const string NUT_ROD3 = " NUT_ROD3 "; // 14
        private const string NUT_TB = " NUT_TB "; // 15
        private const string NUT_NUT5 = " NUT_NUT5 "; // 16
        private const string NUT_NUT6 = " NUT_NUT6 "; // 17
        private const string NUT_ROD4 = " NUT_ROD4 "; // 18
        private const string NUT_TB2 = " NUT_TB2 "; // 19
        private const string NUT_NUT7 = " NUT_NUT7 "; // 20
        private const string NUT_NUT8 = " NUT_NUT8 "; // 21

        String sectionSize, rodSize, beamClampSize, supportType, rodType;
        int topType, numOfRoutes, shoeBegin, shoeEnd, index;
        Double overhang, w1, w2, rodSpacing, supportLength, shoeHeight;
        string[] shoePart;
        Collection<BusinessObject> routeObjects;
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

                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowerAssyWSec", "WSize");
                    int sectionSizeValue = (int)sectionSizeCodelist.PropValue;
                    sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)sectionSizeValue).DisplayName;


                    rodType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrPowerAssySS", "RodType")).PropValue;

                    PropertyValueCodelist topTypeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowerAssySS", "TopType");
                    topType = (int)topTypeCodelist.PropValue;



                    string interfaceName = "";
                    if (support.SupportsInterface("IJUAHgrPowerAssySS4_V"))
                        interfaceName = "IJUAHgrPowerAssySS4_V";
                    else
                        interfaceName = "IJUAHgrPowerAssySS4";

                    PropertyValueCodelist rodSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue(interfaceName, "RodSize");
                    int rodSizeValue = (int)rodSizeCodelist.PropValue;
                    rodSize = rodSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)rodSizeValue).ShortDisplayName;

                    PropertyValueCodelist beamClampSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue(interfaceName, "BeamClampSize");
                    int beamClampSizeValue = (int)beamClampSizeCodelist.PropValue;
                    beamClampSize = beamClampSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)beamClampSizeValue).ShortDisplayName;

                    if (support.SupportsInterface("IJUAHgrPowAssyRodSP_V"))
                        interfaceName = "IJUAHgrPowAssyRodSP_V";
                    else
                        interfaceName = "IJUAHgrPowAssyRodSP";


                    rodSpacing = (double)((PropertyValueDouble)support.GetPropertyValue(interfaceName, "RodSpacing")).PropValue;
                    if (support.SupportsInterface("IJUAHgrPowAssySupLen_V"))
                        interfaceName = "IJUAHgrPowAssySupLen_V";
                    else
                        interfaceName = "IJUAHgrPowAssySupLen";


                    supportLength = (double)((PropertyValueDouble)support.GetPropertyValue(interfaceName, "SupportLength")).PropValue;

                    supportType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrPowAssySSType", "SupType")).PropValue;


                    if (rodSpacing > supportLength)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Rod spacing should be less than the support length. Please check the value.", "", "Assy_SS4_V.cs", 193);
                        return null;
                    }

                    overhang = supportLength / 2 - rodSpacing / 2;
                    w1 = rodSpacing / 2;
                    w2 = rodSpacing / 2;

                    if (topType == 1)//"ROD_CLEVIS_LUG"
                    {
                        if (rodSize == "2")     //For Clevis Lug, we do not haveparts for Rod Size = 2
                            rodSize = "3";


                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(ROD1, rodType + "_" + rodSize));
                        parts.Add(new PartInfo(ROD2, rodType + "_" + rodSize));
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

                    if (topType == 2)//"ROD_BEAM_CLAMP"
                    {

                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(ROD1, rodType + "_" + rodSize));
                        parts.Add(new PartInfo(ROD2, rodType + "_" + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(NUT2, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(BOT_EYE_NUT1, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(BOT_EYE_NUT2, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(BOT_BEAM_ATT1, "Anvil_FIG66_" + rodSize));
                        parts.Add(new PartInfo(BOT_BEAM_ATT2, "Anvil_FIG66_" + rodSize));
                        parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                        parts.Add(new PartInfo(BEAM_CLAMP2, "Anvil_FIG292_" + beamClampSize));
                    }

                    if (topType == 3)//"ROD_BEAM_ATT"
                    {

                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(ROD1, rodType + "_" + rodSize));
                        parts.Add(new PartInfo(ROD2, rodType + "_" + rodSize));
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

                    if (topType == 4)//"ROD_WASHER"
                    {

                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(ROD1, rodType + "_" + rodSize));
                        parts.Add(new PartInfo(ROD2, rodType + "_" + rodSize));
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

                    if (topType == 5)//"ROD_NUT"
                    {

                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(ROD1, rodType + "_" + rodSize));
                        parts.Add(new PartInfo(ROD2, rodType + "_" + rodSize));
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

                    numOfRoutes = SupportHelper.SupportedObjects.Count;
                    NominalDiameter[] pipeDiameter = new NominalDiameter[numOfRoutes];
                    NominalDiameter[] pipeDiameter1 = new NominalDiameter[numOfRoutes];
                    shoePart = new string[numOfRoutes];

                    BusinessObject[] pipe = new BusinessObject[numOfRoutes];
                    routeObjects = new Collection<BusinessObject>();
                    routeObjects = SupportHelper.SupportedObjects;
                    NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000 }, "mm");

                    NominalDiameter minNominalDiameter = new NominalDiameter();
                    minNominalDiameter.Size = 218;
                    minNominalDiameter.Units = "mm";
                    NominalDiameter maxNominalDiameter = new NominalDiameter();
                    maxNominalDiameter.Size = 1020;
                    maxNominalDiameter.Units = "mm";

                    int index = 0;
                    foreach (BusinessObject route in routeObjects)
                    {
                        pipe[index] = route;
                        index++;
                    }

                    for (index = 0; index < numOfRoutes; index++)
                    {


                        if (supportType.ToUpper() == "FIXED")
                        {
                            pipeDiameter[index] = new NominalDiameter();
                            shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssySSShoeH", "ShoeHeight")).PropValue;
                            pipeDiameter[index].Size = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowerAssySSDW", "PipeOutDia")).PropValue;
                            pipeDiameter[index].Units = "mm";

                        }
                        else if (supportType.ToUpper() == "VARIABLE")
                        {

                            PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                            NominalDiameter currentDiameter = new NominalDiameter();
                            currentDiameter = pipeInfo.NominalDiameter;

                            pipeDiameter[index] = new NominalDiameter();
                            pipeDiameter[index].Size = pipeInfo.OutsideDiameter * 1000;
                            pipeDiameter[index].Units = currentDiameter.Units;
                        }
                        pipeDiameter1[index] = new NominalDiameter();
                        if (pipeDiameter[index].Units == "mm")
                        {

                            pipeDiameter1[index].Size = pipeDiameter[index].Size;
                            pipeDiameter1[index].Units = "mm";
                        }
                        else if (pipeDiameter[index].Units == "in")
                        {
                            pipeDiameter1[index].Size = pipeDiameter[index].Size * 39.37008;
                            pipeDiameter1[index].Units = "in";
                        }


                        if (IsPipeSizeValid(pipeDiameter1[index], minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Parts are not available for this Pipe Size.", "", "Assy_SS4_V.cs", 361);
                            return null;
                        }

                        //Get the Shoe    

                        CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                        PartClass auxilaryTable = (PartClass)catalogBaseHelper.GetPartClass("HgrAssy_SZ4PartSel");
                        ReadOnlyCollection<BusinessObject> classItems = auxilaryTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                        foreach (BusinessObject classItem in classItems)
                        {
                            if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAssySZ4PartSel", "PipeOutDia")).PropValue > pipeDiameter1[index].Size - 0.01) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAssySZ4PartSel", "PipeOutDia")).PropValue < pipeDiameter1[index].Size + 0.01) && ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssySZ4PartSel", "UnitType")).PropValue == pipeDiameter1[index].Units))
                            {
                                shoePart[index] = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssySZ4PartSel", "PartNo")).PropValue;

                            }
                        }

                    }
                    shoeBegin = parts.Count + 1;
                    shoeEnd = shoeBegin + numOfRoutes - 1;

                    for (int indexShoe = shoeBegin; indexShoe <= shoeEnd; indexShoe++)
                    {
                        if (indexShoe == shoeBegin)
                            index = 1;
                        else
                            index = indexShoe - shoeBegin + 1;

                        Shoepartkeys[index - 1] = "Shoe" + index.ToString();
                        parts.Add(new PartInfo(Shoepartkeys[index - 1], shoePart[index - 1]));

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
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }

                double BBXWidth = BBX.Width;
                double BBXHeight = BBX.Height;

                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = GetIsLugEndOffsetApplied();

                string[] structPort = new string[2];
                structPort = GetIndexedStructPortName(isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];


                if (IsStructureSlopedAcrossPipe("CONNECTION", false))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "The structure is sloped. Please check.", "", "Assy_SS4.cs", 450);
                    return;
                }
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent horizontalSection = componentDictionary[HOR_SECTION];
                SupportComponent rod1 = componentDictionary[ROD1];
                SupportComponent nut1 = componentDictionary[NUT1];

                SupportComponent[] shoePart = new SupportComponent[numOfRoutes];

                BusinessObject[] shoe = new BusinessObject[numOfRoutes];
                for (int index = 0; index < numOfRoutes; index++)
                {
                    shoePart[index] = componentDictionary[Shoepartkeys[index]];
                    shoe[index] = shoePart[index].GetRelationship("madeFrom", "part").TargetObjects[0];
                }

                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)horizontalSection.GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)horizontalSection.GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;


                horizontalSection.SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                horizontalSection.SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                horizontalSection.SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                horizontalSection.SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                horizontalSection.SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");


                //====== ======
                //Create Joints
                //====== ======


                double horizontalLength1 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", rightStructPort, PortDistanceType.Horizontal);
                double horizontalLength2 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", leftStructPort, PortDistanceType.Horizontal);
                double verticalLength1 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", rightStructPort, PortDistanceType.Vertical);
                double verticalLength2 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", leftStructPort, PortDistanceType.Vertical);
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "BBRV_Low", PortAxisType.X, OrientationAlong.Direct);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if ((w1 + w2) < BBXWidth)
                    {
                        w1 = (BBXWidth + overhang) / 2;
                        w2 = (BBXWidth + overhang) / 2;
                    }

                }
                else
                {
                    if ((w1 + w2) < BBXWidth)
                    {
                        w1 = (BBXWidth + overhang) / 2;
                        w2 = (BBXWidth + overhang) / 2;
                    }
                }

                double CALC1 = Math.Cos(routeStructAngle) * w1;
                double CALC2 = Math.Sin(routeStructAngle) * w1;
                double CALC3 = Math.Cos(routeStructAngle) * w2;
                double CALC4 = Math.Sin(routeStructAngle) * w2;

                Double flangeThickness = 0;
                BusinessObject horizontalSectionPart = horizontalSection.GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                double SteelThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                BusinessObject rod1Part = rod1.GetRelationship("madeFrom", "part").TargetObjects[0];
                BusinessObject nut1Part = nut1.GetRelationship("madeFrom", "part").TargetObjects[0];
                double rodDiameter = (double)((PropertyValueDouble)rod1Part.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue;
                double nutThickness = (double)((PropertyValueDouble)nut1Part.GetPropertyValue("IJUAHgrAnvil_hex_nut", "T")).PropValue;
                //Sort Width (Get the maximum width)
                double[] pipeDiameter = new double[numOfRoutes + 1];
                double[] pipeThickness = new double[numOfRoutes + 1];
                double temp;
                double temp1 = 0;

                for (int index = 0; index <= numOfRoutes; index++)
                {
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    pipeDiameter[index] = pipeInfo.OutsideDiameter;
                    pipeThickness[index] = pipeInfo.InsulationThickness;
                    if (numOfRoutes >= 1)
                    {
                        for (int j = index + 1; j < numOfRoutes; j++)
                        {
                            if (pipeDiameter[j] > pipeDiameter[index])
                            {
                                temp = pipeDiameter[index];
                                pipeDiameter[index] = pipeDiameter[j];
                                pipeDiameter[j] = temp;
                            }
                            if (pipeThickness[j] > pipeThickness[index])
                            {
                                temp = pipeThickness[index];
                                pipeThickness[index] = pipeThickness[j];
                                pipeThickness[j] = temp1;
                            }
                        }
                    }
                }
                double largePipeDiameter = pipeDiameter[0] - 2 * pipeThickness[0];
                Double H = 0;
                if (supportType.ToUpper() == "FIXED")
                {
                    H = shoeHeight;
                    for (index = 0; index < numOfRoutes; index++)
                        shoePart[index].SetPropertyValue(H, "IJUAHgrPowerSZ_ShoeH", "H");

                }
                else if (supportType.ToUpper() == "VARIABLE")
                {

                    H = (double)((PropertyValueDouble)shoe[0].GetPropertyValue("IJUAHgrPowerSZ_ShoeH", "H")).PropValue;
                }


                String structRoutePort;
                for (int indexShoe = shoeBegin; indexShoe <= shoeEnd; indexShoe++)
                {
                    if (indexShoe == shoeBegin)
                    {
                        structRoutePort = "Route";
                    }
                    else
                    {
                        index = indexShoe - shoeBegin + 1;
                        structRoutePort = "Route_" + index;
                    }

                    //Add Joint Between the UBolt and Route


                    JointHelper.CreateRigidJoint(Shoepartkeys[indexShoe - shoeBegin], "Route", "-1", structRoutePort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444
                }
                //Add the intelligence to determine which side of the pipe the steel is when placing by point
                string configuration;
                double beamClampOffset,washerOffset,nutOffset,beamLeftClampByStruct,beamRightClampByStruct;
               

                double byPointAngle1 = GetRouteStructConfigAngle("Route", "Structure", PortAxisType.Y);
                //figure out the orientation of the structure port
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                if (Math.Round(Math.Abs(byPointAngle2), 3) >= Math.Round(Math.PI / 2, 3))   //The structure is oriented in the standard direction
                {
                    if (Math.Round(Math.Abs(byPointAngle1), 3) < Math.Round(Math.PI / 2, 3))
                    {
                        configuration = "1";
                        beamClampOffset = -horizontalLength1 - w1;
                        washerOffset = horizontalLength1 + w1;
                        nutOffset = horizontalLength1 + w1;
                    }
                    else
                    {
                        configuration = "2";
                        beamClampOffset = horizontalLength1 + w1;
                        washerOffset = -horizontalLength1 - w1;
                        nutOffset = horizontalLength1 + w1;
                    }

                }
                else
                {
                    //The structure is oriented in the opposite direction
                    if (Math.Round(Math.Abs(byPointAngle1), 3) < Math.Round(Math.PI / 2, 3))
                    {
                        configuration = "3";
                        beamClampOffset = horizontalLength1 + w1;
                        washerOffset = -horizontalLength1 - w1;
                        nutOffset = horizontalLength1 + w1;
                    }
                    else
                    {
                        configuration = "4";
                        beamClampOffset = -horizontalLength1 - w1;
                        washerOffset = horizontalLength1 + w1;
                        nutOffset = horizontalLength1 + w1;
                    }
                }
                double byPointAngle3 = RefPortHelper.AngleBetweenPorts("BBRV_Low", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Direct);

                double distanceLeftClampRoute = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_Low", PortDistanceType.Horizontal);
                double distanceRightClampRoute = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_High", PortDistanceType.Horizontal);


                if (Math.Abs(byPointAngle3) > Math.PI / 2)    //The structure is oriented in the standard direction
                {
                    beamLeftClampByStruct = distanceLeftClampRoute + w2 - w1 + overhang / 2;
                    beamRightClampByStruct = -distanceRightClampRoute - w2 + w1 - overhang / 2;

                }
                else
                {
                    beamLeftClampByStruct = -distanceLeftClampRoute - w2 + w1 - overhang / 2;
                    beamRightClampByStruct = distanceRightClampRoute + w2 - w1 + overhang / 2;
                }

                //Start Joints
                //Add a Prismatic Joint defining the flexible bottom member
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    horizontalSection.SetPropertyValue(w1 + w2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                else
                {
                    if (leftStructPort == rightStructPort)
                    {
                        if (configuration == "1" || configuration == "3")
                            //Added for TR 111040,104598
                            horizontalSection.SetPropertyValue(w1 + w2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                        else
                            //Added for TR 111040,104598
                            horizontalSection.SetPropertyValue(w1 + w2 + overhang + overhang, "IJUAHgrOccLength", "Length");             //works for beam clamp          
                    }
                    else
                    {
                        horizontalSection.SetPropertyValue(horizontalLength1 + horizontalLength2 + overhang + overhang, "IJUAHgrOccLength", "Length");  //two pieces of steel

                    }
                }

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                //Create the Flexible (Prismatic) Joint between the ports of the top rods
                JointHelper.CreatePrismaticJoint(ROD1, "TopExThdRH", ROD1, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);//3501
                JointHelper.CreatePrismaticJoint(ROD2, "TopExThdRH", ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);//3501

                //Add a Vertical Joint to the Rods Z axes
                JointHelper.CreateGlobalAxesAlignedJoint(ROD1, "TopExThdRH", Axis.Z, Axis.Z);
                JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "TopExThdRH", Axis.Z, Axis.Z);


                if (topType == 1)//ROD_CLEVIS_LUG
                {

                    //Add a revolute Joint between the lug hole and clevis pin
                    JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole", Axis.Y, Axis.X);//74                      
                    JointHelper.CreateRevoluteJoint(CLEVIS2, "Pin", LUG2, "Hole", Axis.Y, Axis.X);//74


                    //Add a rigid Joint between top of the rod and the Clevis
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", CLEVIS, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", CLEVIS2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444
                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, BBXWidth / 2 + w1 + overhang, steelWidth / 2);//1196
                    }
                    else
                    {
                        if (leftStructPort == rightStructPort)   // one piece of steel
                        {
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + BBXWidth / 2 + w1, steelWidth / 2);// 1196
                        }
                        else    // Two pieces of steel
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + horizontalLength1, steelWidth / 2);// 1196

                    }
                    //Add Joints between the lug and the Structure
                    if (Configuration == 1)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, CALC1, -BBXWidth / 2 - CALC2);//10468

                        }
                        else
                        {
                            if (leftStructPort == rightStructPort)   // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, -CALC1 - BBXWidth / 2);//10468

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, -BBXWidth / 2 + CALC1);//

                                else
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, CALC1, -BBXWidth / 2 - CALC2);//

                            }
                            else    // Two pieces of steel
                            {
                                JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength2, 0, horizontalLength2);//

                            }
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, -CALC3, -BBXWidth / 2 + CALC4);//

                        }
                        else
                        {
                            if (leftStructPort == rightStructPort)    // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, CALC3 - BBXWidth / 2);//

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, -BBXWidth / 2 - CALC3);//

                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, -CALC3, -BBXWidth / 2 + CALC4);//

                            }
                            else    //Two pieces of steel
                            {
                                JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, -horizontalLength1);//

                            }
                        }

                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 - CALC2, -CALC1);//

                        else
                        {
                            if (leftStructPort == rightStructPort) // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -CALC1 - BBXWidth / 2, 0);//

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 + CALC1, 0);//

                                else
                                    JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 - CALC2, -CALC1);//

                            }
                            else    // Two pieces of steel
                            {
                                JointHelper.CreateRigidJoint(LUG, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength2, horizontalLength2, 0);//

                            }
                        }


                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 + CALC4, CALC3);//

                        else
                        {
                            if (leftStructPort == rightStructPort)     // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, CALC3 - BBXWidth / 2, 0);//

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 - CALC3, 0);//

                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 + CALC4, CALC3);//

                            }
                            else    //Two pieces of steel
                            {
                                JointHelper.CreateRigidJoint(LUG2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -horizontalLength1, 0);//

                            }
                        }
                    }
                    //Add a Rigid Joint between the bottom eye nut and the rod
                    JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);//10468

                    JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);



                    //Add a Revolute Joint between the bottom eye nut and the beam attachment
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);//81

                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);//81


                    // Add planar joints for bottom Bottom Attachment to horizontal section
                    JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", BOT_BEAM_ATT1, "Structure", Plane.ZX, Plane.XY, 0);


                    JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", BOT_BEAM_ATT2, "Structure", Plane.ZX, Plane.XY, 0);



                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);//

                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);//



                }
                if (topType == 2)//ROD_BEAM_CLAMP
                {
                    SupportComponent beamClamp = componentDictionary[BEAM_CLAMP];
                    SupportComponent beamClamp2 = componentDictionary[BEAM_CLAMP2];

                    double flangeWidth = 0;
                    if (supportingType == "Steel")
                    {
                        flangeWidth = SupportingHelper.SupportingObjectInfo(1).FlangeWidth;
                    }
                    flangeWidth = (Math.Round(flangeWidth) * 1000 / 25.4) + 1;

                    //I know the lowest allowable FLANGE_W is 3 and the max is 15
                    PropertyValueCodelist flangeWidthCodelist = (PropertyValueCodelist)beamClamp.GetPropertyValue("IJOAHgrAnvil_FIG292", "FLANGE_WIDTH");
                    flangeWidthCodelist.PropValue = (int)flangeWidth;
                    if (flangeWidthCodelist.PropValue < 3)
                        flangeWidthCodelist.PropValue = 3;
                    else
                        if (flangeWidthCodelist.PropValue > 15)
                            flangeWidthCodelist.PropValue = 15;




                    beamClamp.SetPropertyValue(flangeWidthCodelist.PropValue, "IJOAHgrAnvil_FIG292", "FLANGE_WIDTH");
                    beamClamp2.SetPropertyValue(flangeWidthCodelist.PropValue, "IJOAHgrAnvil_FIG292", "FLANGE_WIDTH");


                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, BBXWidth / 2 + w1 + overhang, steelWidth / 2);//1196

                    else
                    {
                        if (leftStructPort == rightStructPort)    // one piece of steel
                        {
                            if (configuration == "1" || configuration == "3")
                                JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, BBXWidth / 2 + w1 + overhang, steelWidth / 2);//1196

                            else
                                JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + BBXWidth / 2 + w1, steelWidth / 2);//1196


                        }
                        else    // Two pieces of steel
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + horizontalLength1, steelDepth / 2);//1196


                    }



                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, w1);//beamLeftClampByStruct);//9380

                    else
                    {
                        if (leftStructPort == rightStructPort)    // one piece of steel
                        {
                            if (configuration == "1" || configuration == "2")
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, w1);//9380

                            else
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, w1, 0);//9380

                        }
                        else    //Two pieces of steel
                            JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);


                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, -w2);// beamRightClampByStruct);                        
                    else
                        if (supportingType == "Slab")
                            JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", leftStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -w2, 0);
                        else
                            JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", leftStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, -w2);


                    //Add a Planar Joint between top of the rod and the Beam Clamp
                    JointHelper.CreatePlanarJoint(ROD1, "TopExThdRH", BEAM_CLAMP, "InThdRH", Plane.XY, Plane.XY, 0);//100

                    JointHelper.CreatePlanarJoint(ROD2, "TopExThdRH", BEAM_CLAMP2, "InThdRH", Plane.XY, Plane.XY, 0);//100



                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);//81

                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);//81



                    //Add a Rigid Joint between the beam attachment and the route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC2, -CALC1);//9380

                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC4, CALC3);//9380


                    }
                    else
                    {
                        if (leftStructPort == rightStructPort)    // one piece of steel
                        {
                            if (configuration == "3" || configuration == "4")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC1, 0);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC3, 0);//9380


                            }
                            else if (configuration == "1")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC1, 0);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC3, 0);//9380


                            }
                            else
                            {
                                //Changed this for TR111040,104598
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC2, -CALC1);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC4, CALC3);//9380


                            }
                        }
                        else    // Two pieces of steel
                        {
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, horizontalLength1, 0);//9380

                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, -horizontalLength2, 0);//9380


                        }
                    }

                    //Add rigid joint between bottom rods and eye nuts
                    JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444

                    JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444


                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);//9444

                    JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);//9444



                }
                if (topType == 3)//ROD_BEAM_ATT
                {

                    //Add a revolute Joint between the eye nut and beam att pin
                    JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.Y, Axis.X);//74    

                    JointHelper.CreateRevoluteJoint(EYE_NUT2, "Eye", BEAM_ATT2, "Pin", Axis.Y, Axis.X);//74


                    //Add a rigid Joint between top of the rod and the Clevis
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444

                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", EYE_NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444







                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, BBXWidth / 2 + w1 + overhang, steelWidth / 2);//1196

                    else
                    {
                        if (leftStructPort == rightStructPort)    // one piece of steel
                        {
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + BBXWidth / 2 + w1, steelWidth / 2);
                        }
                        else    // Two pieces of steel
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + horizontalLength1, steelWidth / 2);

                    }


                    //Add a rigid joint between the beam attachment and the route
                    if (Configuration == 1)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 + CALC4, CALC3);

                        else
                        {
                            if (leftStructPort == rightStructPort)    // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 + CALC3, 0);

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 - CALC3, 0);

                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 + CALC4, CALC3);


                            }
                            else    // Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -horizontalLength1, 0);


                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 - CALC2, -CALC1);

                        else
                        {
                            if (leftStructPort == rightStructPort)    // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 - CALC1, 0);

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 + CALC1, 0);

                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength1, -BBXWidth / 2 - CALC2, -CALC1);


                            }
                            else    // Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -verticalLength2, horizontalLength2, 0);

                        }

                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, -CALC3, -BBXWidth / 2 + CALC4);//10468

                        else
                        {
                            if (leftStructPort == rightStructPort)    // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, -BBXWidth / 2 - CALC3);

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, -BBXWidth / 2 + CALC3);

                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, -CALC3, -BBXWidth / 2 + CALC4);

                            }
                            else    // Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, 0, -horizontalLength1);

                        }


                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, CALC1, -BBXWidth / 2 - CALC2);

                        else
                        {
                            if (leftStructPort == rightStructPort)    // one piece of steel
                            {
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, BBXWidth / 2 - CALC1, 0);

                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, -BBXWidth / 2 - CALC1, 0);

                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength1, CALC1, -BBXWidth / 2 - CALC2);

                            }
                            else    // Two pieces of steel
                                JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, -verticalLength2, 0, horizontalLength2);


                        }
                    }

                    // Add a Planar Joint between the beam attachment and horizontal section
                    JointHelper.CreatePlanarJoint(BOT_BEAM_ATT1, "Structure", HOR_SECTION, "BeginCap", Plane.XY, Plane.ZX, 0);//108

                    JointHelper.CreatePlanarJoint(BOT_BEAM_ATT2, "Structure", HOR_SECTION, "EndCap", Plane.XY, Plane.ZX, 0);


                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);//81

                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);//81




                    //Add a Rigid Joint between the bottom eye nut and the rod
                    JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444

                    JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444


                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);

                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);

                }




                if (topType == 4)//ROD_WASHER
                {

                    flangeThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                    //Make revolute joint between rods and the connection object
                    JointHelper.CreatePlanarJoint(ROD1, "TopExThdRH", CONNECTION, "Connection", Plane.XY, Plane.NegativeXY, 0);//36

                    JointHelper.CreatePlanarJoint(ROD2, "TopExThdRH", CONNECTION2, "Connection", Plane.XY, Plane.NegativeXY, 0);



                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, BBXWidth / 2 + w1 + overhang, steelWidth / 2);//1196

                    else
                    {
                        if (leftStructPort == rightStructPort) // one piece of steel
                        {
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + BBXWidth / 2 + w1, steelWidth / 2);
                        }
                        else    // Two pieces of steel
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + horizontalLength1, steelWidth / 2);


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
                        //TR90541
                        if (leftStructPort == rightStructPort)    // one piece of steel                
                            JointHelper.CreateRigidJoint(WASHER, "Structure", ROD2, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);

                        else    // Two pieces of steel
                            JointHelper.CreateRigidJoint(WASHER, "Structure", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0.01, 0, 0);

                    }

                    //Add a Prismatic Joint between the lug and the Structure
                    //TR90541
                    JointHelper.CreateRigidJoint(WASHER2, "Structure", ROD1, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);


                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);//81

                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);//81


                    //Add a Rigid Joint between the beam attachment and the route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC2, -CALC1);//9380

                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC4, CALC3);//9380


                    }
                    else
                    {
                        if (leftStructPort == rightStructPort)    // one piece of steel
                        {
                            if (configuration == "3" || configuration == "4")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC1, 0);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC3, 0);//9380


                            }
                            else if (configuration == "1")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC1, 0);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC3, 0);//9380


                            }
                            else
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC2, -CALC1);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC4, CALC3);//9380


                            }

                        }
                        else    // Two pieces of steel
                        {
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, horizontalLength1, 0);//9380

                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, -horizontalLength2, 0);//9380


                        }
                    }




                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", WASH_NUT9, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);

                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", WASH_NUT10, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);

                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", WASH_NUT11, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);

                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", WASH_NUT12, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);



                    // Add rigid joint between bottom rods and eye nuts
                    JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444

                    JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);//9444




                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);

                    JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);



                }

                if (topType == 5)//ROD_NUT
                {
                    // dFlangeThickness = PW_GetSupportingPropertyByRule("tf", 1, myIJHgrInputConfigHlpr)
                    flangeThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                    //Make revolute joint between rods and the connection object
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
                        if (leftStructPort == rightStructPort)  //one piece of steel
                        {
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, nutOffset, 0);

                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);


                        }
                        else    // Two pieces of steel
                        {
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);

                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05, 0, 0);


                        }
                    }
                    //Joint from W-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, BBXWidth / 2 + w1 + overhang, steelWidth / 2);

                    else
                    {
                        if (leftStructPort == rightStructPort)   // one piece of steel
                        {
                            if (configuration == "1" || configuration == "3")
                                //TR 111040
                                JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + BBXWidth / 2 + w1, steelWidth / 2);

                            else
                                JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + BBXWidth / 2 + w1, steelDepth / 2);

                        }
                        else    // Two pieces of steel
                            JointHelper.CreateRigidJoint("-1", "BBRV_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -H + largePipeDiameter / 2, overhang + horizontalLength1, steelDepth / 2);


                    }


                    //Add a Rigid Joint between the beam attachment and the route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC2, -CALC1);//9380

                        JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC4, CALC3);//9380


                    }
                    else
                    {
                        if (leftStructPort == rightStructPort)   // one piece of steel
                        {
                            if (configuration == "1")
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC1, 0);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC3, 0);//9380
                            }
                            else if (configuration == "3" || configuration == "4")
                            {
                                if (HgrCompareDoubleService.cmpdbl(byPointAngle3, 0) == true)
                                {
                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC1, 0);//9380

                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC3, 0);//9380

                                }
                                else
                                {
                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC1, 0);//9380

                                    JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC3, 0);//9380

                                }
                            }
                            else
                            {
                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 + CALC2, -CALC1);//9380

                                JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, BBXWidth / 2 - CALC4, CALC3);//9380

                            }
                        }
                        else    // Two pieces of steel
                        {
                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT1, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, horizontalLength1, 0);//9380

                            JointHelper.CreateRigidJoint(BOT_BEAM_ATT2, "Structure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -H + largePipeDiameter / 2, -horizontalLength2, 0);//9380


                        }
                    }

                    //Add a revolute Joint between the bottom nut and the beam attachments
                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT1, "Eye", BOT_BEAM_ATT1, "Pin", Axis.X, Axis.Y);//81

                    JointHelper.CreateRevoluteJoint(BOT_EYE_NUT2, "Eye", BOT_BEAM_ATT2, "Pin", Axis.X, Axis.Y);//81




                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT9, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);

                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT10, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);

                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT_NUT11, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);

                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT_NUT12, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);



                    //Add rigid joint between bottom rods and eye nuts
                    JointHelper.CreateRigidJoint(BOT_EYE_NUT1, "InThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    JointHelper.CreateRigidJoint(BOT_EYE_NUT2, "InThdRH", ROD2, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);



                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);

                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -rodDiameter * 2 - nutThickness, 0, 0);


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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(HOR_SECTION, 1));
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
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(HOR_SECTION, 1));
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

        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string BOMString = "";
            try
            {

                NominalDiameter pipeDiameter = new NominalDiameter();
                Collection<BusinessObject> pipes = SupportHelper.SupportedObjects;


                string bomDescription = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJOAHgrPowAssyBOMDesc", "BOM_DESC")).PropValue;

                supportType = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAHgrPowAssySSType", "SupType")).PropValue;
                if (supportType.ToUpper() == "FIXED")
                {
                    pipeDiameter.Size = (double)((PropertyValueDouble)SupportOrComponent.GetPropertyValue("IJUAHgrPowerAssySSDW", "PipeOutDia")).PropValue;

                }
                else if (supportType.ToUpper() == "VARIABLE")
                {

                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    NominalDiameter currentDiameter = new NominalDiameter();
                    currentDiameter = pipeInfo.NominalDiameter;

                    pipeDiameter.Size = pipeInfo.OutsideDiameter;
                    pipeDiameter.Units = currentDiameter.Units;
                    double insulationThickness = pipeInfo.InsulationThickness;
                    pipeDiameter.Size = pipeDiameter.Size - 2 * insulationThickness;
                    if (pipeDiameter.Units == "mm")
                    {

                        pipeDiameter.Size = pipeDiameter.Size * 1000;
                        pipeDiameter.Units = "mm";
                    }
                    else if (pipeDiameter.Units == "in")
                    {
                        pipeDiameter.Size = pipeDiameter.Size * 39.37008;
                        pipeDiameter.Units = "in";
                    }
                }


                if (bomDescription == null)
                {
                    if (supportType.ToUpper() == "FIXED")
                        BOMString = "SS4 Assembly for " + Math.Round(pipeDiameter.Size, 1) + " mm Pipe Outside Dia with fixed overall width and rod spacing";
                    else if (supportType.ToUpper() == "VARIABLE")
                        BOMString = "SS4 Assembly for " + Math.Round(pipeDiameter.Size, 1) + " " + pipeDiameter.Units + " with variable overall width and rod spacing";
                }
                else
                {
                    BOMString = bomDescription;
                }



                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Assy_ss4_V" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="IsOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public String[] GetIndexedStructPortName(Boolean[] IsOffsetApplied)
        {
            String[] structurePort = new String[2];
            int structureCount = SupportHelper.SupportingObjects.Count;
            int i;
            string supportingType = String.Empty;

            if ((SupportHelper.SupportingObjects.Count != 0))
            {
                if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                    supportingType = "Steel";    //Steel                      
                else
                    supportingType = "Slab"; //Slab
            }
            else
                supportingType = "Slab"; //For PlaceByReference

            if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
            {
                structurePort[0] = "Structure";
                structurePort[1] = "Structure";
            }
            else
            {
                structurePort[0] = "Structure";
                structurePort[1] = "Struct_2";

                if (structureCount > 1)
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        for (i = 0; i <= 1; i++)
                        {
                            double angle = 0;
                            if ((supportingType == "Steel") && IsOffsetApplied[i] == false)
                            {
                                angle = RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                            }
                            //the port is the right structure port
                            if (Math.Abs(angle) < Math.PI / 2.0)
                            {
                                if (i == 0)
                                {
                                    structurePort[0] = "Struct_2";
                                    structurePort[1] = "Structure";
                                }
                            }
                            //the port is the left structure port
                            else
                            {
                                if (i == 1)
                                {
                                    structurePort[0] = "Struct_2";
                                    structurePort[1] = "Structure";
                                }
                            }
                        }
                    }
                }
                else
                    structurePort[1] = "Structure";
            }
            //switch the OffsetApplied flag
            if (structurePort[0] == "Struct_2")
            {
                Boolean flag = IsOffsetApplied[0];
                IsOffsetApplied[0] = IsOffsetApplied[1];
                IsOffsetApplied[1] = flag;
            }

            return structurePort;
        }
        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public Boolean[] GetIsLugEndOffsetApplied()
        {
            try
            {
                Collection<BusinessObject> StructureObjects;
                Boolean[] isOffsetApplied = new Boolean[2];

                //first route object is set as primary route object
                StructureObjects = SupportHelper.SupportingObjects;
                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;

                if (StructureObjects != null)
                {
                    if (StructureObjects.Count >= 2)
                    {
                        for (int index = 0; index <= 1; index++)
                        {
                            //check if offset is to be applied
                            const double RouteStructAngle = Math.PI / 180.0;
                            Boolean varRuleApplied = true;

                            if (supportingType == "Steel")
                            {
                                //if angle is within 1 degree, regard as parallel case
                                //Also check for Sloped structure                                
                                MemberPart memberPart = (MemberPart)SupportHelper.SupportingObjects[index];
                                ICurve memberCurve = memberPart.Axis;

                                Vector supportedVector = new Vector();
                                Vector supportingVector = new Vector();

                                if (SupportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                                {
                                    Position startLocation = new Position(SupportedHelper.SupportedObjectInfo(1).StartLocation);
                                    Position endLocation = new Position(SupportedHelper.SupportedObjectInfo(1).EndLocation);
                                    supportedVector = new Vector(endLocation - startLocation);
                                }
                                if (memberCurve is ILine)
                                {
                                    ILine line = (ILine)memberCurve;
                                    supportingVector = line.Direction;
                                }

                                double angle = GetAngleBetweenVectors(supportingVector, supportedVector);
                                double refAngle1 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - Math.PI / 2;
                                double refAngle2 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                if (angle < (refAngle1 + 0.001) && angle > (refAngle1 - 0.001))
                                    angle = angle - Math.Abs(refAngle1);
                                else if (angle < (refAngle2 + 0.001) && angle > (refAngle2 - 0.001))
                                    angle = angle - Math.Abs(refAngle2);
                                else
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                if (Math.Abs(angle) < RouteStructAngle || Math.Abs(angle - Math.PI) < RouteStructAngle)
                                    varRuleApplied = false;
                            }

                            isOffsetApplied[index] = varRuleApplied;
                        }
                    }
                }

                return isOffsetApplied;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied Method of Bline_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        public bool IsStructureSlopedAcrossPipe(String structType, Boolean onTurn)
        {


            Double refAngle, refAngle2;
            refAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Y, OrientationAlong.Global_Z);
            refAngle2 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z);
            bool isStructureSlopedAcrossPipe = false;

            if (structType == "Slab")
            {
                if (refAngle > ((89.99999999999) * Math.PI / 180.0) && refAngle < ((90.0000000001) * Math.PI / 180.0))
                    isStructureSlopedAcrossPipe = false;
                else
                    isStructureSlopedAcrossPipe = true;

            }
            else
            {
                if (refAngle2 > ((89.99999999999) * Math.PI / 180.0) && refAngle2 < ((90.0000000001) * Math.PI / 180.0))
                    isStructureSlopedAcrossPipe = false;
                else
                {
                    if ((SupportHelper.PlacementType == PlacementType.PlaceByStruct) || onTurn)
                        isStructureSlopedAcrossPipe = true;
                    else
                        isStructureSlopedAcrossPipe = false;
                }
            }
            return isStructureSlopedAcrossPipe;
        }
        /// <summary>
        /// This method returns the direct angle between Route and structur ports in Radians
        /// </summary>
        /// <param name="routePortName">string-route Port Name.</param>
        /// <param name="structPortName">string-structure Port Name.</param>
        /// <param name="axisType">PortAxisType-axis Type.</param>
        /// <returns>double</returns>        
        /// <code>
        ///   byPointAngle1=GetRouteStructConfigAngle("Route", "Structure", PortAxisType.Y);
        ///</code>
        public double GetRouteStructConfigAngle(String routePortName, String structPortName, PortAxisType axisType)
        {
            try
            {

                //get the appropriate axis
                Vector[] vecAxis = new Vector[2];

                Matrix4X4 routeMatrix = RefPortHelper.PortLCS(routePortName);
                Position routepoint = routeMatrix.Origin;

                switch (axisType)
                {
                    case PortAxisType.X:
                        {
                            vecAxis[0] = routeMatrix.XAxis;
                            break;
                        }
                    case PortAxisType.Y:
                        {
                            vecAxis[0] = routeMatrix.ZAxis.Cross(routeMatrix.XAxis);
                            break;
                        }
                    case PortAxisType.Z:
                        {
                            vecAxis[0] = routeMatrix.ZAxis;
                            break;
                        }
                }
                Matrix4X4 structMatrix = RefPortHelper.PortLCS(structPortName);
                Position structPoint = structMatrix.Origin;
                vecAxis[1] = structPoint.Subtract(routepoint);

                return GetAngleBetweenVectors(vecAxis[0], vecAxis[1]);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///       ContentHelper contentHelper = new ContentHelper();
        ///       double value;
        ///       value = contentHelper. GetAngleBetweenVectors(vector1, vector2 );
        ///</code>

        public double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dblDotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double dblArcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dblDotProd), 1) == false)
                {
                    dblArcCos = Math.PI / 2 - Math.Atan(dblDotProd / Math.Sqrt(1 - dblDotProd * dblDotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dblDotProd, -1) == true)
                {
                    dblArcCos = Math.PI;
                }
                else if (HgrCompareDoubleService.cmpdbl(dblDotProd, 1) == true)
                {
                    dblArcCos = 0;
                }
                return dblArcCos;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}