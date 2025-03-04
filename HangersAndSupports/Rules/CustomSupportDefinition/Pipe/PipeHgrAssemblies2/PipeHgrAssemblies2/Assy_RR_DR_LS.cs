//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_DR_LS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_DR_LS
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01-10-2015      Vinay   DI-CP-276996	Update HS_Assembly2 to use RichSteel 
//   17/12/2015      Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   21/03/2016      Vinay   TR-CP-288920	Issues found in HS_Assembly_V2
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

    public class Assy_RR_DR_LS : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_DR_LS"
        //----------------------------------------------------------------------------------

        //Constants
        //For everything
        private const string HOR_SECTION = "hor_Section";
        private const string ROD1 = "rod1";
        private const string ROD2 = "rod2";
        private const string NUT1 = "nut1";
        private const string NUT2 = "nut2";
        private const string NUT3 = "nut3";
        private const string NUT4 = "nut4";

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

        private string rodType = string.Empty;
        private string topType = string.Empty;

        private double W1;
        private double W2;
        private double overhang;
        private double bottomRodLength;
        private int turnbuckle = 0;
        int index = 0;
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
                    string sectionSize = ((PropertyValueString)support.GetPropertyValue("IJUAHgrLSize", "LSize")).PropValue;
                    rodType = ((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRodType", "ROD_TYPE")).PropValue;
                    topType = ((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyTopType", "TOP_TYPE")).PropValue;

                    PropertyValueCodelist turnbuckleCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTurnbuckle", "TURNBUCKLE");
                    turnbuckle = turnbuckleCodelist.PropValue;

                    overhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrTrapeze", "OVERHANG")).PropValue;
                    bottomRodLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyRR", "BOT_ROD_LENGTH")).PropValue;

                    W1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W1")).PropValue;
                    W2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_RR_DR", "W2")).PropValue;

                    if (topType.ToUpper() == "ROD_CLEVIS_LUG")
                    {
                        if (turnbuckle == 1)//1 means With Turnbuckle
                        {
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                            parts.Add(new PartInfo(ROD1, rodType));
                            parts.Add(new PartInfo(ROD2, rodType));
                            parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                            parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
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
                return 4;
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
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BoundingBox boundingBox = null;
                //==========================
                //Custom bounding box definition
                //==========================
                //Create Vectors to define the plane of the BBX
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                if (Configuration == 3)
                {
                    //BBSR
                    Vector globalZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); //Get Global Z
                    Vector boundingBox_X = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, globalZ); //Project Route into Horizontal Plane
                    Vector boundingBox_Z = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, boundingBox_X); //Project Vector From Route to Structure into the BBX Plane
                    BoundingBoxHelper.CreateBoundingBox(boundingBox_Z, boundingBox_X,"RRAssyBBX", false, true, true, false);
                }
                else if (Configuration == 4)
                {
                    //BBRV
                    Vector globalZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); //Get Global Z
                    Vector boundingBox_X = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, globalZ); //Project Route into Horizontal Plane
                    Vector boundingBox_Z = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, boundingBox_X); //Project Vector From Route to Structure into the BBX Plane
                    Vector boundingBox_Y = globalZ.Cross(boundingBox_X); //Get Horizontal Vector in the BBX Plane (For Orthogonal direction)

                    //Get Orthogonal Vector in the plane of the BBX (Gz, -Gz, By, -By) depending on Angle 
                    if (Math.Acos(globalZ.Dot(boundingBox_Z) / (globalZ.Length * boundingBox_Z.Length)) < Math.Round(Math.PI / 4, 11))
                        boundingBox_Z = globalZ;
                    else if (Math.Acos(globalZ.Dot(boundingBox_Z) / (globalZ.Length * boundingBox_Z.Length)) > 3 * Math.Round(Math.PI / 4, 11))
                        boundingBox_Z.Set(-globalZ.X, -globalZ.Y, globalZ.Z);
                    else if (Math.Acos(boundingBox_Y.Dot(boundingBox_Z) / (boundingBox_Y.Length * boundingBox_Z.Length)) > 3 * Math.Round(Math.PI / 4, 11))
                        boundingBox_Z = boundingBox_Y;
                    else
                        boundingBox_Z.Set(-boundingBox_Y.X, -boundingBox_Y.Y, boundingBox_Y.Z);

                    BoundingBoxHelper.CreateBoundingBox(boundingBox_Z, boundingBox_X, "RRAssyBBX", false, true, true, false);
                }

                string portBBSR = string.Empty;
                string portBBRV = string.Empty;
                if (Configuration == 3 || Configuration == 4)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        portBBRV = "RRAssyBBX_Low";
                        portBBSR = "RRAssyBBX_Low";
                    }
                    else
                    {
                        portBBSR = "RRAssyBBX_Low";
                        portBBRV = "RRAssyBBX_Low";
                    }
                }
                else if (Configuration == 1 || Configuration == 2)
                {
                    portBBSR = "BBSR_Low";
                    portBBRV = "BBRV_Low";
                }
                //====== ======
                //3. retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry

                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | dHeight
                // |____________________|
                //        dWidth

                Double width, height;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

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
                //Create Joints
                //====== ======

                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                double lengthHor1, lengthHor2, lengthVert1, lengthVert2;
                if (topType.ToUpper() == "ROD_BEAM_CLAMP")
                {
                    lengthHor1 = RefPortHelper.DistanceBetweenPorts(portBBRV, rightStructPort, PortDistanceType.Horizontal);
                    lengthHor2 = RefPortHelper.DistanceBetweenPorts(portBBRV, leftStructPort, PortDistanceType.Horizontal);
                    lengthVert1 = RefPortHelper.DistanceBetweenPorts(portBBRV, rightStructPort, PortDistanceType.Vertical);
                    lengthVert2 = RefPortHelper.DistanceBetweenPorts(portBBRV, leftStructPort, PortDistanceType.Vertical);
                }
                else
                {
                    lengthHor1 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", rightStructPort, PortDistanceType.Horizontal);
                    lengthHor2 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", leftStructPort, PortDistanceType.Horizontal);
                    lengthVert1 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", rightStructPort, PortDistanceType.Vertical);
                    lengthVert2 = RefPortHelper.DistanceBetweenPorts("BBRV_Low", leftStructPort, PortDistanceType.Vertical);
                }

                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "BBRV_Low", PortAxisType.X, OrientationAlong.Direct);

                //Can't find a way to look up the actual PIPE DIAMETER without including the insulation thickness
                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;
                if ((W1 + W2) < width)
                {
                    W1 = (width + overhang) / 2;
                    W2 = (width + overhang) / 2;
                }

                support.SetPropertyValue(W1, "IJOAHgrAssy_RR_DR", "W1");
                support.SetPropertyValue(W2, "IJOAHgrAssy_RR_DR", "W2");

                double calc1 = Math.Cos(routeStructAngle) * W1;
                double clac2 = Math.Sin(routeStructAngle) * W1;
                double calc3 = Math.Cos(routeStructAngle) * W2;
                double calc4 = Math.Sin(routeStructAngle) * W2;

                double flangeThickness = 0;
                BusinessObject horizontalSectionPart = componentDictionary[HOR_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                double steelThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                BusinessObject rod1 = componentDictionary[ROD1].GetRelationship("madeFrom", "part").TargetObjects[0];
                double rodDiameter = (double)((PropertyValueDouble)rod1.GetPropertyValue("IJUAhsRodDiameter", "RodDiameter")).PropValue;
                BusinessObject nut1 = componentDictionary[NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                double nutT = (double)((PropertyValueDouble)nut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                //figure out the orientation of the structure port
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);


                string configuration = string.Empty;
                double beamClampOffset, beamLeftClampByStruct, beamRightClampByStruct;
                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI / 2, 11))    //The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 11))
                    {
                        configuration = "1";
                        beamClampOffset = -lengthHor1 - W1;
                    }
                    else
                    {
                        configuration = "2";
                        beamClampOffset = lengthHor1 + W1;
                    }
                }
                else    //The structure is oriented in the opposite direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 11))
                    {
                        configuration = "3";
                        beamClampOffset = lengthHor1 + W1;
                    }
                    else
                    {
                        configuration = "4";
                        beamClampOffset = -lengthHor1 - W1;
                    }
                }

                double byPointAngle3 = RefPortHelper.AngleBetweenPorts("BBRV_Low", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Direct);
                double distLeftClampRoute = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_Low", PortDistanceType.Horizontal);
                double distRightClampRoute = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_High", PortDistanceType.Horizontal);
                if (SupportHelper.SupportedObjects.Count == 1)
                {
                    if (Math.Abs(byPointAngle3) > Math.Round(Math.PI / 2, 7))    //The structure is oriented in the standard direction
                    {
                        beamLeftClampByStruct = W1;
                        beamRightClampByStruct = -W2;
                    }
                    else
                    {
                        beamLeftClampByStruct = -W1;
                        beamRightClampByStruct = W2;
                    }
                }
                else
                {
                    if ((Math.Abs(byPointAngle3) >= Math.Round(Math.PI / 2, 7)) && (Math.Abs(byPointAngle3) <= Math.Round(Math.PI / 2, 7)))//The structure is oriented in the standard direction
                    {
                        beamLeftClampByStruct = distLeftClampRoute + W2 - W1 + overhang / 2;
                        beamRightClampByStruct = -distRightClampRoute - W2 + W1 - overhang / 2;
                    }
                    else
                    {
                        beamLeftClampByStruct = -distLeftClampRoute - W2 + W1 - overhang / 2;
                        beamRightClampByStruct = distRightClampRoute + W2 - W1 + overhang / 2;
                    }
                }

                //Start the Joints here ********************************************************************************************************************
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    componentDictionary[HOR_SECTION].SetPropertyValue(W1 + W2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                else
                {
                    if (leftStructPort == rightStructPort)
                        componentDictionary[HOR_SECTION].SetPropertyValue(W1 + W2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                    else
                        componentDictionary[HOR_SECTION].SetPropertyValue(lengthHor1 + lengthHor2 + overhang + overhang, "IJUAHgrOccLength", "Length");
                }
                //**************************************************************************************************
                //*****************************************
                // L Note
                //***************************************** 
                CodelistItem fabrication;
                ControlPoint controlPoint;
                Note note;
                double endNotePosition = (double)((PropertyValueDouble)(componentDictionary[HOR_SECTION].GetPropertyValue("IJUAHgrOccLength", "Length"))).PropValue;
                string[] noteName = new string[] { "L_Start", "L_Mid", "L_End", "L_1", "L_2" };
                note = CreateNote("L_Start", HOR_SECTION, "BeginCap", new Position(0, 0 , 0), " ", true, 2, 1, out controlPoint);
                note = CreateNote("L_Mid", HOR_SECTION, "BeginCap", new Position(0, 0, overhang), " ", true, 2, 1, out controlPoint);
                note = CreateNote("L_End", HOR_SECTION, "BeginCap", new Position(0, 0, endNotePosition), " ", true, 2, 1, out controlPoint);
                if (Configuration == 1 || Configuration == 2)
                {
                    note = CreateNote("L_1", HOR_SECTION, "BeginCap", new Position(0, -lengthVert1, overhang), " ", true, 2, 1, out controlPoint);
                    note = CreateNote("L_2", HOR_SECTION, "BeginCap", new Position(0, -lengthVert2, endNotePosition - overhang), " ", true, 2, 1, out controlPoint);
                }
                else
                {
                    note = CreateNote("L_1", HOR_SECTION, "BeginCap", new Position(0, -lengthVert2, overhang), " ", true, 2, 1, out controlPoint);
                    note = CreateNote("L_2", HOR_SECTION, "BeginCap", new Position(0, -lengthVert1, endNotePosition - overhang), " ", true, 2, 1, out controlPoint);
                }


                for (index = 0; index < 5; index++)
                {
                    note.SetPropertyValue(noteName[index], "IJGeneralNote", "Text");
                    fabrication = note.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    note.SetPropertyValue(fabrication, "IJGeneralNote", "Purpose"); //value 3 means fabrication                   
                    note.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                
                //**************************************************************************************************
                //Add a Vertical Joint to the Rods Z axes
                JointHelper.CreateGlobalAxesAlignedJoint(ROD1, "RodEnd1", Axis.Z, Axis.Z);
                JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "RodEnd1", Axis.Z, Axis.Z);
                //Create the Flexible (Prismatic) Joint between the ports of the top rods
                JointHelper.CreatePrismaticJoint(ROD1, "RodEnd1", ROD1, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                JointHelper.CreatePrismaticJoint(ROD2, "RodEnd1", ROD2, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                if (topType.ToUpper() == "ROD_CLEVIS_LUG")
                {
                    //Add a revolute Joint between the lug hole and clevis pin
                    JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole1", Axis.Y, Axis.Y);
                    JointHelper.CreateRevoluteJoint(CLEVIS2, "Pin", LUG2, "Hole1", Axis.Y, Axis.Y);

                    //Add a rigid Joint between top of the rod and the Clevis
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CLEVIS, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", CLEVIS2, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Joint from L-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", portBBSR, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + W1 + overhang, steelDepth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort)
                            if (configuration == "1" || configuration == "3")
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelWidth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelDepth / 2);
                        else
                            if (Configuration == 1 || Configuration == 2)
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelDepth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor2 + width, steelDepth / 2);
                    }
                    //Add Joints between the lug and the Structure
                     if (Configuration == 1 || Configuration == 3)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - clac2, -calc1);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -calc1 - width / 2, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - clac2, -calc1);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert2, lengthHor2, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, lengthHor1-width, 0);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc4, calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, calc3 - width / 2, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -width / 2 + calc4, calc3);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert1, -lengthHor1, 0);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.X, -lengthVert2, -lengthHor2-width, 0);
                        }
                        
                    }
                    else if (Configuration == 2 || Configuration == 4)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - clac2);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -calc1 - width / 2);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 + calc1);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, calc1, -width / 2 - clac2);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert2, 0, lengthHor2);
                                else
                                    JointHelper.CreateRigidJoint(LUG, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, lengthHor1 - width);
                        }
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, calc3 - width / 2);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -width / 2 - calc3);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, -calc3, -width / 2 + calc4);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert1, 0, -lengthHor1);
                                else
                                    JointHelper.CreateRigidJoint(LUG2, "Hole2", "-1", portBBRV, Plane.XY, Plane.XY, Axis.X, Axis.Y, -lengthVert2, 0, -lengthHor2 - width);
                        }
                    }
                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[LUG_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[LUG_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD4, "RodEnd1", Axis.Z, Axis.Z);
                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", LUG_ROD3, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", LUG_ROD4, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
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
                        JointHelper.CreateRigidJoint(LUG_ROD3, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT*3, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD4, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT*3, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD4, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[LUG_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = (opening1 - rodTakeOut1) / 2 + shapeLength1;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[LUG_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = (opening2 - rodTakeOut2) / 2 + shapeLength2;


                        JointHelper.CreateRigidJoint(LUG_ROD3, "RodEnd2", LUG_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LUG_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_ROD4, "RodEnd2", LUG_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", LUG_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                    }
                    else
                    {
                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION,"BeginCap", ROD1, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", ROD2, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT*3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT*3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                    }
                }
                if (topType.ToUpper() == "ROD_BEAM_CLAMP")
                {
                    //Joint from L-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct )
                        JointHelper.CreateRigidJoint("-1", portBBSR, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + W1 + overhang, steelDepth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort)
                            if (configuration == "1" || configuration == "3")
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + W1 + overhang, steelWidth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelDepth / 2);
                        else
                        {
                            if (Configuration==1 || Configuration==2)
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelDepth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor2, steelDepth / 2);
                        }
                    }
                    //Connect Beam clamps to the structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct )
                        JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -beamLeftClampByStruct);
                    else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 - calc3, 0);
                    }
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            if (configuration == "1" || configuration == "2")
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, W1, 0);
                        else //Two pieces of steel
                            JointHelper.CreateRigidJoint(BEAM_CLAMP, "Structure", "-1", rightStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -beamRightClampByStruct);
                    else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 + calc1, 0);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            JointHelper.CreateRigidJoint(BEAM_CLAMP2, "Structure", "-1", leftStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -W2, 0);
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
                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ROD3, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", ROD4, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(ROD4, "RodEnd1", Axis.Z, Axis.Z);
                        //Add a rigid Joint between bottom rods and Turnbuckles
                        JointHelper.CreateRigidJoint(ROD3, "RodEnd2", TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "RodEnd2", TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(ROD3, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                         BusinessObject turnBucklePart1 = componentDictionary[TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = (opening1 - rodTakeOut1) / 2 + shapeLength1;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = (opening2 - rodTakeOut2) / 2 + shapeLength2;

                        JointHelper.CreateRigidJoint(ROD3, "RodEnd2", NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD4, "RodEnd2", NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                    }
                    else
                    {
                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ROD1, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", ROD2, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT*3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT*3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
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

                    //Joint from L-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", portBBSR, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + W1 + overhang, steelDepth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort)
                            if (configuration == "1" || configuration == "3")
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelWidth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelDepth / 2);
                        else
                            if (Configuration == 1 || Configuration == 2)
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelDepth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0,  width + lengthHor2 + overhang, steelDepth / 2);
                    }
                    //Add Joints between the beam attachment and the Structure 
                    if (Configuration == 1 || Configuration == 3)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 - calc4, calc3);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 - calc3, 0);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 + calc3, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, -width / 2 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 - clac2, 0);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, lengthHor1, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert2, lengthHor2 + width, 0);

                        }
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 + clac2, -calc1);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 + calc1, 0);//width / 2 + calc1
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, width / 2 + calc4, 0);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert2, -lengthHor2, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1, -lengthHor1 + width, 0);
                        }
                    }
                    else if (Configuration == 2 || Configuration == 4)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1, calc3, -width / 2 + calc4);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                             JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 - calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 - calc3);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 + calc3);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, -calc3, -width / 2 + calc4);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, 0, -lengthHor1);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert2, 0, -lengthHor2 - width);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, -calc1, -width / 2 - clac2);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 + calc1);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, 0, -width / 2 + calc1);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, -width / 2 - calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, calc1, -width / 2 - clac2);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert2, 0, lengthHor2);
                                else
                                    JointHelper.CreateRigidJoint(BEAM_ATT2, "Structure", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1, 0, lengthHor1 - width);
                        }
                    }
                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[ATT_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[ATT_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        //Add planar joints for bottom rods to horizontal section

                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ATT_ROD3, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", ATT_ROD4, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);

                        //Add a Vertical Joint to the Rods Z axes
                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD4, "RodEnd1", Axis.Z, Axis.Z);

                        //Add a Joint between bottom rod and Turnbuckle
                        JointHelper.CreateRigidJoint(ATT_ROD3, "RodEnd2", ATT_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "RodEnd2", ATT_TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a Joint between top rod and Turnbuckle
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
                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(ATT_ROD3, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[ATT_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = (opening1 - rodTakeOut1) / 2 + shapeLength1;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[ATT_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = (opening2 - rodTakeOut2) / 2 + shapeLength2;

                        JointHelper.CreateRigidJoint(ATT_ROD3, "RodEnd2", ATT_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", ATT_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ATT_ROD4, "RodEnd2", ATT_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", ATT_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                    }
                    else
                    {
                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ROD1, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ROD2, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                    }
                }
                if (topType.ToUpper() == "ROD_WASHER")
                {
                    //We need to get the Flange Width.  However, it comes in metres, so we change it to inches.  Then we Int it to get rid of decimals.
                    //Then we add one so that the Flange Width of the clamp is bigger than the actual flange width, not smaller

                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeThickness = 0.02;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                                flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                        }
                        else
                            flangeThickness = 0.02;
                    }

                    //Add a Rigid Joint between top of the rod and the Beam Clamp
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", CONNECTION2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Joint from L-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", portBBSR, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + W1 + overhang, steelDepth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelDepth / 2);
                        else
                            if (Configuration == 1 || Configuration == 2)
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelDepth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor2+width, steelDepth / 2);
                    }

                    //Add a Rigid Joint between the connection and the Route
                    if (Configuration == 1 || Configuration == 3)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - calc4, calc3);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - calc3, 0);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + calc3, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, -width / 2 - calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - clac2, 0);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, lengthHor1, 0);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert2 + 0.05 + flangeThickness / 2, lengthHor2 + width, 0);

                        }
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + clac2, -calc1);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + calc1, 0);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, -width / 2 + calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, width / 2 + calc4, 0);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert2 + 0.05 + flangeThickness / 2, -lengthHor2, 0);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, lengthVert1 + 0.05 + flangeThickness / 2, -lengthHor1 + width, 0);
                        }
                    }
                    else if (Configuration == 2 || Configuration == 4)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, calc3, -width / 2 + calc4);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 - calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc3);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc3);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc4);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -lengthHor1);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert2 + 0.05 + flangeThickness / 2, 0, -lengthHor2 - width);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, -calc1, -width / 2 - clac2);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 + calc1);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 - calc1);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 - calc1);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, -width / 2 - clac2);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert2 + 0.05 + flangeThickness / 2, 0, lengthHor2);
                                else
                                    JointHelper.CreateRigidJoint(CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, lengthVert1 + 0.05 + flangeThickness / 2, 0, lengthHor1 - width);
                        }
                    }
                    //Add a Prismatic Joint between the lug and the Structure
                    JointHelper.CreateRigidJoint(WASHER, "Port2", ROD2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 - 0.05, 0, 0);

                    //Add a Prismatic Joint between the lug and the Structure
                    JointHelper.CreateRigidJoint(WASHER2, "Port2", ROD1, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 - 0.05, 0, 0);

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", WASH_NUT9, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.03 + nutT, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", WASH_NUT10, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.015 + nutT, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", WASH_NUT11, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.03 + nutT, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", WASH_NUT12, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.015 + nutT, 0, 0);
                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[WASH_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[WASH_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");

                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", WASH_ROD3, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", WASH_ROD4, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);

                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD4, "RodEnd1", Axis.Z, Axis.Z);
                        //Add a rigid Joint between bottom rods and Turnbuckles

                        JointHelper.CreateRigidJoint(WASH_ROD3, "RodEnd2", WASH_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "RodEnd2", WASH_TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add a rigid Joint between top rods and Turnbuckle
                        JointHelper.CreateRigidJoint(WASH_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(WASH_ROD3, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[WASH_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = (opening1 - rodTakeOut1) / 2 + shapeLength1;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[WASH_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = (opening2 - rodTakeOut2) / 2 + shapeLength2;

                        JointHelper.CreateRigidJoint(WASH_ROD3, "RodEnd2", WASH_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", WASH_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(WASH_ROD4, "RodEnd2", WASH_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", WASH_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                    }
                    else
                    {
                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ROD1, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ROD2, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                    }
                }
                if (topType.ToUpper() == "ROD_NUT")
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeThickness = 0.02;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                            flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                        else
                            flangeThickness = 0.02;
                    }
                    //Add a Rigid Joint between top of the rod and the Beam Clamp
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

                    //Add a Rigid Joint between Connection  and Route
                    if (Configuration == 1 || Configuration == 3)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 - calc4, calc3);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 - calc3, 0);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, -width / 2 + calc3, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 + calc3, 0);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 + clac2, 0);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, lengthHor1, 0);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, lengthHor2 + width, 0);

                        }
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 + clac2, -calc1);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 + calc1, 0);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, -width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 - calc1, 0);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, width / 2 - calc4, 0);
                            else //Two pieces of steel
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, -lengthHor2, 0);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.X, topNutPosition, -lengthHor1 + width, 0);
                        }
                    }
                    else if (Configuration == 2 || Configuration == 4)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, topNutPosition, calc3, -width / 2 + calc4);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 - calc3);
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 - calc3);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 - calc3);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, -calc3, -width / 2 + calc4);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, -lengthHor1);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, -lengthHor2 - width);
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, -calc1, -width / 2 - clac2);
                        else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            if (SupportHelper.SupportedObjects.Count > 1)
                                JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, width / 2 - calc1 + overhang);
                            else
                                JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0,  calc1 - overhang);
                        }
                        else
                        {
                            if (leftStructPort == rightStructPort) //one piece of steel
                                if (configuration == "3" || configuration == "4")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, width / 2 - calc1, 0);
                                else if (configuration == "1")
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, -width / 2 + calc1);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, calc1, -width / 2 - clac2);
                            else //Two pieces of steel
                                if (Configuration == 2)
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, lengthHor2);
                                else
                                    JointHelper.CreateRigidJoint(NUT_CONNECTION2, "Connection", "-1", portBBRV, Plane.NegativeXY, Plane.XY, Axis.X, Axis.Y, topNutPosition, 0, lengthHor1 - width);
                        }
                    }

                    //Joint from L-Section to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", portBBSR, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, width / 2 + W1 + overhang, steelDepth / 2);
                    else
                    {
                        if (leftStructPort == rightStructPort) //one piece of steel
                            if (configuration == "1" || configuration == "3")
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelWidth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + width / 2 + W1, steelDepth / 2);
                        else
                            JointHelper.CreateRigidJoint("-1", portBBRV, HOR_SECTION, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, 0, overhang + lengthHor1, steelDepth / 2);
                    }
                    //Add a joint between the top nuts and the top rods
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT_NUT9, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05 + nutT, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT_NUT10, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 + 0.05, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", NUT_NUT11, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05 + nutT, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "RodEnd2", NUT_NUT12, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 + 0.05, 0, 0);
                    if (turnbuckle == 1)//1 means With Turnbuckle
                    {
                        componentDictionary[NUT_ROD3].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        componentDictionary[NUT_ROD4].SetPropertyValue(bottomRodLength, "IJOAHgrOccLength", "Length");
                        // Add a rigid joint between the horizontal section and the bottom rods
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", NUT_ROD3, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", NUT_ROD4, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        //Add a Vertical Joint to the Rods Z axes

                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD3, "RodEnd1", Axis.Z, Axis.Z);
                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD4, "RodEnd1", Axis.Z, Axis.Z);
                        //Add a rigid Joint between bottom rods and Turnbuckles

                        JointHelper.CreateRigidJoint(NUT_ROD3, "RodEnd2", NUT_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "RodEnd2", NUT_TB2, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a rigid Joint between top rods and Turnbuckle

                        JointHelper.CreateRigidJoint(NUT_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_TB2, "RodEnd1", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add a Rigid Joint between the nuts and the rods
                        JointHelper.CreateRigidJoint(NUT_ROD3, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD3, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);

                        //Getting the nut positions for Turnbuckle1
                        BusinessObject turnBucklePart1 = componentDictionary[NUT_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength1 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition1 = (opening1 - rodTakeOut1) / 2 + shapeLength1;

                        //Getting the nut positions for Turnbuckle2
                        BusinessObject turnBucklePart2 = componentDictionary[NUT_TB2].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength2 = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition2 = (opening2 - rodTakeOut2) / 2 + shapeLength2;

                        JointHelper.CreateRigidJoint(NUT_ROD3, "RodEnd2", NUT_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT_NUT6, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(NUT_ROD4, "RodEnd2", NUT_NUT7, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition2 + nutT, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT_NUT8, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutT, 0, 0);

                    }
                    else
                    {
                        //Add planar joints for bottom rods to horizontal section
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "BeginCap", ROD1, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3);
                        JointHelper.CreatePlanarJoint(HOR_SECTION, "EndCap", ROD2, "RodEnd1", Plane.ZX, Plane.NegativeXY, steelThickness + nutT * 3); ;
                        //Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 3, 0, 0);
                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutT * 2, 0, 0);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    for (int iIndex = 1; iIndex <= SupportHelper.SupportedObjects.Count; iIndex++)
                        routeConnections.Add(new ConnectionInfo(HOR_SECTION, iIndex));

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
                    int supportingObjects;
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        supportingObjects = 1;
                    else
                        supportingObjects = SupportHelper.SupportingObjects.Count;


                    if (topType == "ROD_CLEVIS_LUG")
                    {
                        structConnections.Add(new ConnectionInfo(LUG, 1));
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
                if (eMirrorPlane == MirrorPlane.YZPlane)
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 3;
                    else if (CurrentMirrorToggleValue == 2)
                        return 4;
                    else if (CurrentMirrorToggleValue == 3)
                        return 1;
                    else
                        return 2;
                }
                else
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

