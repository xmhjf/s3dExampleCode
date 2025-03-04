//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_LR.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_LR
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
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

    public class Assy_RR_LR : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_LR"
        //----------------------------------------------------------------------------------

        //Constants

        private const string CLAMP_TYPE = "CLAMP_TYPE";
        private const string DIM_1 = "DIM_1";
        private const string DIM_2 = "DIM_2";
        private const string BOT_EYENUT_1 = "BOT_EYENUT_1";
        private const string BOT_EYENUT_2 = "BOT_EYENUT_2";
        private const string HEX_NUT = "HEX_NUT";
        private const string HEX_NUT_1 = "HEX_NUT_1";
        private const string HEX_NUT_2 = "HEX_NUT_2";
        private const string HEX_NUT_3 = "HEX_NUT_3";
        private const string WBA_1 = "WBA_1";
        private const string MID_ROD = "MID_ROD";
        private const string TURNBUCKLE_1 = "TURNBUCKLE_1";
        private const string COUPLING = "COUPLING";
        private const string WBA_TYPE_1 = "WBA_TYPE_1";
        private const string EYE_NUT_1 = "EYE_NUT_1";
        private const string CLEVIS_1 = "CLEVIS_1";
        private const string LUG_1 = "LUG_1";
        private const string WPLATE_1 = "WPLATE_1";
        private const string BOT_ROD_1 = "BOT_ROD_1";

        private const double CONST_1 = 3.6576;
        private const double CONST_PIPEDIAMETER = 0.1143;


        double clamp_to, rodDiameter = 0, hexNutWeight = 0, length = 0, rodDiameterOverride, eyeNutTo = 0, beamAttachmentTo = 0, clevisTo = 0, turnbuckleTo = 0;
        double lugTo = 0, angularMovementX, angularMovementY, maxAngularMovement, midRodLength, shortRodLength, bottomRodLength, pipeDiameter, channelDistance;
        string turnbuckleFlag = string.Empty, trigger = string.Empty, supportingPartKey;
        private int midrod_counter, hexnut_counter, coupling_counter, noOfRods;

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    double clampABMax = 0, clampCMax = 0, clampDMax = 0, globalAMFM = 0, globalCFM = 0, globalDFM = 0, clampWeight = 0;
                    double rodDiameterAB = 0, rodDiameterC = 0, rodDiameterD = 0, eyeNutWeight = 0, turnbuckleWeight = 0, midRodWeight = 0;
                    string clampType = string.Empty, bottomEyeNutType = string.Empty;
                    double botRodWeight = 0, inchRodDiameter = 0, beamAttachmentWidth = 0, lugRodSize = 0, lugT = 0, wplateW = 0, clevisWeight = 0;
                    double couplingWeight = 0, effectiveLength, totalRodLength, noOfCouplers, clampPIn = 0;

                    hexnut_counter = 0; midrod_counter = 0; coupling_counter = 0;

                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                    pipeDiameter = pipeInfo.OutsideDiameter;

                    length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                    double thermal_X = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "THERMAL_X")).PropValue;
                    double thermal_Y = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "THERMAL_Y")).PropValue;
                    double thermal_Z = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "THERMAL_Z")).PropValue;
                    double abFx_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "ABFx_M")).PropValue;
                    double abFy_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "ABFy_M")).PropValue;
                    double abFz_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "ABFz_M")).PropValue;
                    double abFx_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "ABFx_P")).PropValue;
                    double abFy_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "ABFy_P")).PropValue;
                    double abFz_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "ABFz_P")).PropValue;
                    double cFx_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "CFx_M")).PropValue;
                    double cFy_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "CFy_M")).PropValue;
                    double cFz_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "CFz_M")).PropValue;
                    double cFx_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "CFx_P")).PropValue;
                    double cFy_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "CFy_P")).PropValue;
                    double cFz_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "CFz_P")).PropValue;
                    double dFx_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "DFx_M")).PropValue;
                    double dFy_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "DFy_M")).PropValue;
                    double dFz_M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "DFz_M")).PropValue;
                    double dFx_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "DFx_P")).PropValue;
                    double dFy_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "DFy_P")).PropValue;
                    double dFz_P = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "DFz_P")).PropValue;

                    double insulT = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "INSUL_T")).PropValue;
                    PropertyValueCodelist adjustableCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "ADJUSTABLE");
                    string adjustable = adjustableCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(adjustableCodeList.PropValue).DisplayName;
                    
                    globalAMFM = abFz_M;    // get the value entered by user into a global varible
                    globalCFM = cFz_M;    // get the value entered by user into a global varible
                    globalDFM = dFz_M;    // get the value entered by user into a global varible
                    //====== ======
                    // Set Values of Part Occurance Attributes
                    //====== ======
                    double backToBack;
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        backToBack = 0.02;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && SupportingHelper.SupportingObjectInfo(1).SectionType == "2L")
                                backToBack = SupportingHelper.SupportingObjectInfo(1).Gap;
                            else
                                backToBack = 0;
                        }
                        else
                            backToBack = 0;
                    }
                    support.SetPropertyValue(backToBack, "IJUAHgrAssy_RR_LR", "CHANNEL_DIS");
                    channelDistance = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssy_RR_LR", "CHANNEL_DIS")).PropValue;

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    ReadOnlyCollection<BusinessObject> fig140NParts, fig290NParts, hexNutParts, fig230NParts, fig66NParts, fig299NParts, fig55NSParts, fig60NParts, fig135NParts, fig253NParts, fig295NParts, fig295HNParts, fig212NParts, fig216NParts, clampTypeParts;

                    //*******************************************************************************************************************************************************************
                    //Determine Pipe Clamp Type
                    if (insulT > 0.1016)//NEED TO MAKE THIS TO METRIC
                        throw new CmnException("Insulation Thickness cannot be larger than 4in.");
                    
                    if (insulT > 0)
                    {
                        if (pipeDiameter >= 0.033401 && pipeDiameter <= 0.9144)
                        {
                            PartClass lrParts_fig295NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG295N");
                            fig295NParts = lrParts_fig295NPartClass.Parts;
                            foreach (BusinessObject part in fig295NParts)
                            {
                                if (((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295N", "PipeDiam")).PropValue >= pipeDiameter - 0.001) && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295N", "PipeDiam")).PropValue <= (pipeDiameter + 0.001)))
                                {
                                    clampABMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295N", "AB_MAX")).PropValue;
                                    clampCMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295N", "C_MAX")).PropValue;
                                    clampDMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295N", "D_MAX")).PropValue;
                                    break;
                                }
                            }
                            if (abFz_M > clampABMax || cFz_M > clampCMax || dFz_M > clampDMax)
                            {
                                if (pipeDiameter >= 0.033401 && pipeDiameter <= 0.9144)
                                {

                                    PartClass lrParts_FIG295HNPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG295HN");
                                    fig295HNParts = lrParts_FIG295HNPartClass.Parts;
                                    foreach (BusinessObject part in fig295HNParts)
                                    {
                                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295HN", "PipeDiam")).PropValue >= pipeDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295HN", "PipeDiam")).PropValue <= pipeDiameter + 0.001)
                                        {
                                            clampABMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295HN", "AB_MAX")).PropValue;
                                            clampCMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295HN", "C_MAX")).PropValue;
                                            clampDMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG295HN", "D_MAX")).PropValue;
                                            break;
                                        }
                                    }
                                    if (abFz_M > clampABMax || cFz_M > clampCMax || dFz_M > clampDMax)
                                        throw new CmnException("No suitable double-bolt clamp available.");
                                    else
                                        clampType = "LRParts_FIG295HN";
                                }
                                else
                                    throw new CmnException("No suitable double-bolt clamp available.");
                            }
                            else
                                clampType = "LRParts_FIG295N";
                        }
                        else
                            throw new CmnException("No suitable double-bolt clamp available.");

                    }
                    else
                    {
                        if (pipeDiameter >= 0.02667 && pipeDiameter <= 0.762)
                        {
                            PartClass lrParts_fig212NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG212N");
                            fig212NParts = lrParts_fig212NPartClass.Parts;
                            foreach (BusinessObject part in fig212NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG212N", "PipeDiam")).PropValue >= pipeDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG212N", "PipeDiam")).PropValue <= pipeDiameter + 0.001)
                                {
                                    clampABMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG212N", "AB_MAX")).PropValue;
                                    clampCMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG212N", "C_MAX")).PropValue;
                                    clampDMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG212N", "D_MAX")).PropValue;
                                    break;
                                }
                            }
                            if (abFz_M > clampABMax || cFz_M > clampCMax || dFz_M > clampDMax)
                            {
                                if (pipeDiameter >= 0.02667 && pipeDiameter <= 0.762)
                                {
                                    PartClass LRParts_FIG216NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG216N");
                                    fig216NParts = LRParts_FIG216NPartClass.Parts;
                                    foreach (BusinessObject part in fig216NParts)
                                    {
                                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG216N", "PipeDiam")).PropValue >= pipeDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG216N", "PipeDiam")).PropValue <= pipeDiameter + 0.001)
                                        {
                                            clampABMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG216N", "AB_MAX")).PropValue;
                                            clampCMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG216N", "C_MAX")).PropValue;
                                            clampDMax = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG216N", "D_MAX")).PropValue;
                                            break;
                                        }
                                    }
                                    if (abFz_M > clampABMax || cFz_M > clampCMax || dFz_M > clampDMax)
                                        throw new CmnException("No suitable clamp available.");
                                    else
                                        clampType = "LRParts_FIG216N";
                                }
                                else
                                    throw new CmnException("No suitable clamp available.");
                            }
                            else
                                clampType = "LRParts_FIG212N";
                        }
                        else
                            throw new CmnException("No suitable clamp available.");
                    }
                    //*******************************************************************************************************************************************************************
                    string tempTableName = string.Empty;
                    tempTableName = "IJUAHgr" + clampType;
                    PartClass clampTypePartClass = (PartClass)catalogBaseHelper.GetPartClass(clampType);
                    clampTypeParts = clampTypePartClass.Parts;
                    foreach (BusinessObject part in clampTypeParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue(tempTableName, "PipeDiam")).PropValue >= pipeDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue(tempTableName, "PipeDiam")).PropValue <= pipeDiameter + 0.001)
                        {
                            clamp_to = (double)((PropertyValueDouble)part.GetPropertyValue(tempTableName, "TAKE_OUT")).PropValue;
                            clampPIn = (double)((PropertyValueDouble)part.GetPropertyValue(tempTableName, "F")).PropValue;
                            clampWeight = (double)((PropertyValueDouble)part.GetPropertyValue(tempTableName, "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    if (clamp_to < 0)
                        throw new CmnException("Selected Pipe Size is not available with Clamp.");
                    //********************************************************************************************************************************************************************
                    //Select Rod Size of Eye Nut based on loads
                    abFz_M = globalAMFM + clampWeight;
                    cFz_M = globalCFM + clampWeight;
                    dFz_M = globalDFM + clampWeight;

                    PartClass lrParts_fig290NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG290N");
                    fig290NParts = lrParts_fig290NPartClass.Parts;
                    foreach (BusinessObject part in fig290NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "AB_MIN")).PropValue <= abFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "AB_MAX")).PropValue >= abFz_M)
                            rodDiameterAB = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue;

                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "C_MIN")).PropValue <= cFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "C_MAX")).PropValue >= cFz_M)
                            rodDiameterC = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue;

                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "D_MAX")).PropValue >= dFz_M)
                            rodDiameterD = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue;

                    }
                    if (rodDiameterAB < 0 || rodDiameterC < 0 || rodDiameterD < 0)
                        throw new CmnException("Loads are not within recommended range.");

                    if (rodDiameterAB > rodDiameterC)
                    {
                        if (rodDiameterAB > rodDiameterD)
                        {
                            foreach (BusinessObject part in fig290NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "AB_MIN")).PropValue <= abFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "AB_MAX")).PropValue >= abFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (BusinessObject part in fig290NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "D_MAX")).PropValue >= dFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (rodDiameterC > rodDiameterD)
                        {
                            foreach (BusinessObject part in fig290NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "C_MIN")).PropValue <= cFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "C_MAX")).PropValue >= cFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (BusinessObject part in fig290NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "D_MAX")).PropValue >= dFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                    }
                    foreach (BusinessObject part in fig290NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            eyeNutWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    PartClass lrpartsHexNutPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_HEX_NUT");
                    hexNutParts = lrpartsHexNutPartClass.Parts;
                    foreach (BusinessObject part in hexNutParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            hexNutWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    if (pipeDiameter > CONST_PIPEDIAMETER)
                        hexNutWeight = hexNutWeight * 2;

                    //********************************************************************************************************************************************************************

                    //********************************************************************************************************************************************************************
                    //Select Rod Size of Turnbuckle based on loads
                    abFz_M = globalAMFM + clampWeight + eyeNutWeight;
                    cFz_M = globalCFM + clampWeight + eyeNutWeight;
                    dFz_M = globalDFM + clampWeight + eyeNutWeight;

                    PartClass lrParts_fig230NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG230N");
                    fig230NParts = lrParts_fig230NPartClass.Parts;
                    foreach (BusinessObject part in fig230NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "AB_MIN")).PropValue <= abFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "AB_MAX")).PropValue >= abFz_M)
                            rodDiameterAB = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue;

                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "C_MIN")).PropValue <= cFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "C_MAX")).PropValue >= cFz_M)
                            rodDiameterC = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue;

                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "D_MAX")).PropValue >= dFz_M)
                            rodDiameterD = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue;

                    }
                    if (rodDiameterAB < 0 || rodDiameterC < 0 || rodDiameterD < 0)
                        throw new CmnException("Loads are not within recommended range.");
                    if (rodDiameterAB > rodDiameterC)
                    {
                        if (rodDiameterAB > rodDiameterD)
                        {
                            foreach (BusinessObject part in fig230NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "AB_MIN")).PropValue <= abFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "AB_MAX")).PropValue >= abFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (BusinessObject part in fig230NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "D_MAX")).PropValue >= dFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (rodDiameterC > rodDiameterD)
                        {
                            foreach (BusinessObject part in fig230NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "C_MIN")).PropValue <= cFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "C_MAX")).PropValue >= cFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (BusinessObject part in fig230NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "D_MAX")).PropValue >= dFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                    }
                    foreach (BusinessObject part in fig230NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            turnbuckleWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    if ((adjustable == "No") || (adjustable == "Auto" && ((length - clamp_to * 2.0) < 0.8382)))    //Approximate Rod Length
                        turnbuckleWeight = 0;
                    //********************************************************************************************************************************************************************

                    //***************************************************************************************
                    //Calculate the Weight of the Rod for Single Rod Assemblies
                    abFz_M = globalAMFM + eyeNutWeight + clampWeight + hexNutWeight + turnbuckleWeight;
                    cFz_M = globalCFM + eyeNutWeight + clampWeight + hexNutWeight + turnbuckleWeight;
                    dFz_M = globalDFM + eyeNutWeight + clampWeight + hexNutWeight + turnbuckleWeight;

                    PartClass lrParts_fig140NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG140N");
                    fig140NParts = lrParts_fig140NPartClass.Parts;
                    foreach (BusinessObject part in fig140NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "AB_MIN")).PropValue <= abFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "AB_MAX")).PropValue >= abFz_M)
                            rodDiameterAB = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "C_MIN")).PropValue <= cFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "C_MAX")).PropValue >= cFz_M)
                            rodDiameterC = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MAX")).PropValue >= dFz_M)
                            rodDiameterD = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                    }
                    if (rodDiameterAB < 0 || rodDiameterC < 0 || rodDiameterD < 0)
                        throw new CmnException("Loads are not within recommended range.");
                    if (rodDiameterAB > rodDiameterC)
                    {
                        if (rodDiameterAB > rodDiameterD)
                        {
                            foreach (BusinessObject part in fig140NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "AB_MIN")).PropValue <= abFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "AB_MAX")).PropValue >= abFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (BusinessObject part in fig140NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MAX")).PropValue >= dFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (rodDiameterC > rodDiameterD)
                        {
                            foreach (BusinessObject part in fig140NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "C_MIN")).PropValue <= cFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "C_MAX")).PropValue >= cFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (BusinessObject part in fig140NParts)
                            {
                                if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MAX")).PropValue >= dFz_M)
                                {
                                    rodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                                    break;
                                }
                            }
                        }
                    }
                    if (pipeDiameter < 0.060325 && rodDiameter < 0.009525)
                        rodDiameter = 0.009525;
                    if (pipeDiameter > 0.0635 && rodDiameter < 0.0127)
                        rodDiameter = 0.0127;
                    foreach (BusinessObject part in fig140NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            midRodWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "WEIGHT_PER_LENGTH")).PropValue * (length - clamp_to) / 0.3048;    //Approximate Rod Length
                            break;
                        }
                    }
                    //***************************************************************************************

                    //********************************************************************************************************************************************************************
                    //Get WEIGHT of EYE NUT based on ROD_SIZE
                    foreach (BusinessObject part in fig290NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            eyeNutWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    //********************************************************************************************************************************************************************

                    //********************************************************************************************************************************************************************
                    //Get WEIGHT of Turnbuckle based on ROD_SIZE
                    foreach (BusinessObject part in fig230NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            turnbuckleWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    if ((adjustable == "No") || (adjustable == "Auto" && ((length - clamp_to * 2.0) < 0.8382)))    //Approximate Rod Length
                        turnbuckleWeight = 0;
                    //********************************************************************************************************************************************************************

                    //********************************************************************************************************************************************************************
                    //Get WEIGHT of HEX NUT based on ROD_SIZE
                    foreach (BusinessObject part in hexNutParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            hexNutWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "WEIGHT")).PropValue;
                            if (pipeDiameter > CONST_PIPEDIAMETER)
                            {
                                hexNutWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "WEIGHT")).PropValue * 2;
                                break;
                            }
                        }
                    }
                    //********************************************************************************************************************************************************************

                    //***************************************************************************************
                    //Re-Calculate the Weight of the Rod for Single Rod Assemblies
                    abFz_M = globalAMFM + eyeNutWeight + clampWeight + hexNutWeight + turnbuckleWeight;
                    cFz_M = globalCFM + eyeNutWeight + clampWeight + hexNutWeight + turnbuckleWeight;
                    dFz_M = globalDFM + eyeNutWeight + clampWeight + hexNutWeight + turnbuckleWeight;

                    foreach (BusinessObject part in fig140NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "AB_MIN")).PropValue <= abFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "AB_MAX")).PropValue >= abFz_M)
                            rodDiameterAB = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "C_MIN")).PropValue <= cFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "C_MAX")).PropValue >= cFz_M)
                            rodDiameterC = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MIN")).PropValue <= dFz_M && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "D_MAX")).PropValue >= dFz_M)
                            rodDiameterD = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue;
                    }
                    if (rodDiameterAB < 0 || rodDiameterC < 0 || rodDiameterD < 0)
                        throw new CmnException("Loads are not within recommended range.");
                    if (rodDiameterAB > rodDiameterC)
                    {
                        if (rodDiameterAB > rodDiameterD)
                            rodDiameter = rodDiameterAB;
                        else
                            rodDiameter = rodDiameterD;
                    }
                    else
                    {
                        if (rodDiameterC > rodDiameterD)
                            rodDiameter = rodDiameterC;
                        else
                            rodDiameter = rodDiameterD;
                    }
                    if (pipeDiameter < 0.060325 && rodDiameter < 0.009525)
                        rodDiameter = 0.009525;
                    if (pipeDiameter > 0.0635 && rodDiameter < 0.0127)
                        rodDiameter = 0.0127;
                    foreach (BusinessObject part in fig140NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            midRodWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "WEIGHT_PER_LENGTH")).PropValue * (length - clamp_to) / 0.3048;    //Approximate Rod Length
                            break;
                        }
                    }
                    //***************************************************************************************


                    //************************************************************************************************************************************************************************
                    rodDiameterOverride = rodDiameter;

                    //*************************** Get final takeouts of everything
                    foreach (BusinessObject part in fig290NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            eyeNutTo = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "E")).PropValue - rodDiameter / 2.0;
                            break;
                        }
                    }

                    PartClass lrParts_fig66NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG66N");
                    fig66NParts = lrParts_fig66NPartClass.Parts;
                    foreach (BusinessObject part in fig66NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig66N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig66N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            beamAttachmentTo = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig66N", "E_PRIME")).PropValue - (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig66N", "H")).PropValue / 2.0;
                            beamAttachmentWidth = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig66N", "S")).PropValue - (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig66N", "T")).PropValue / 2.0;
                            break;
                        }
                    }
                    PartClass lrParts_fig299NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG299N");
                    fig299NParts = lrParts_fig299NPartClass.Parts;
                    foreach (BusinessObject part in fig299NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            clevisTo = -(double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "ROD_DIA")).PropValue / 2.0 - (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "P")).PropValue / 2.0 + (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "A")).PropValue;
                            break;
                        }
                    }
                    turnbuckleTo = 0.0762;
                    //Welding Lug doesn't have a rod size of 3/8in
                    if (HgrCompareDoubleService.cmpdbl(rodDiameterOverride , 0.009525)==true)
                        lugRodSize = 0.0127;
                    else
                        lugRodSize = rodDiameterOverride;

                    PartClass lrParts_fig55NSPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG55NS");
                    fig55NSParts = lrParts_fig55NSPartClass.Parts;
                    foreach (BusinessObject part in fig55NSParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG55NS", "ROD_DIA")).PropValue >= lugRodSize - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG55NS", "ROD_DIA")).PropValue <= lugRodSize + 0.001)
                        {
                            lugTo = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG55NS", "H")).PropValue + (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG55NS", "F")).PropValue / 2.0;
                            lugT = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG55NS", "T")).PropValue;
                            break;
                        }
                    }
                    PartClass lrParts_fig60NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG60N");
                    fig60NParts = lrParts_fig60NPartClass.Parts;
                    foreach (BusinessObject part in fig60NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG60N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG60N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            wplateW = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG60N", "WIDTH")).PropValue;
                            break;
                        }
                    }

                    //*****************************************************************************************
                    //Rules
                    if (HgrCompareDoubleService.cmpdbl(abFx_M, 0) == false || HgrCompareDoubleService.cmpdbl(cFx_M, 0) == false || HgrCompareDoubleService.cmpdbl(dFx_M, 0) == false || HgrCompareDoubleService.cmpdbl(abFy_M, 0) == false || HgrCompareDoubleService.cmpdbl(cFy_M, 0) == false || HgrCompareDoubleService.cmpdbl(dFy_M, 0) == false)
                        throw new CmnException("Loads are not within recommended range.");
                    //*****************************************************************************************

                    //************************************************************************************************************************************************************************
                    //TRACK DOWN THE FINAL WEIGHTS OF EVERYTHING
                    foreach (BusinessObject part in fig290NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            eyeNutWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG290N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    foreach (BusinessObject part in hexNutParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            hexNutWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_hex_nut", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    foreach (BusinessObject part in fig230NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue >= rodDiameter - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "ROD_DIA")).PropValue <= rodDiameter + 0.001)
                        {
                            turnbuckleWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG230N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    foreach (BusinessObject part in fig299NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            clevisWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_fig299N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    PartClass lrParts_FIG135NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG135N");
                    fig135NParts = lrParts_FIG135NPartClass.Parts;
                    foreach (BusinessObject part in fig135NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG135N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG135N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            couplingWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG135N", "WEIGHT")).PropValue;
                            break;
                        }
                    }
                    //'************************************************************************************************************************************************************************

                    //'*********************************************************************************************************
                    //'Check the Side Movement
                    effectiveLength = length - clamp_to + clampPIn / 2.0;
                    totalRodLength = effectiveLength - eyeNutTo * 2.0 - beamAttachmentTo - clampPIn / 2.0;
                    angularMovementX = Math.Atan(thermal_X / effectiveLength);
                    angularMovementY = Math.Atan(thermal_Y / effectiveLength);
                    if (Math.Abs(angularMovementX) > 5 * (3.14 / 180))
                    {
                        //If want to raise error set value for last argument as "False".
                        //If it is a warning then set the value to "True".
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Side movement too high for use of Pipe Clamp. Choose another support method..", "", "Assy_RR_LR.cs", 1);
                        return null;
                    }
                    
                    if (Math.Abs(angularMovementX) > Math.Abs(angularMovementY))
                        maxAngularMovement = Math.Abs(angularMovementX);
                    else
                        maxAngularMovement = Math.Abs(angularMovementY);
                    //*********************************************************************************************************

                    //Determine Rod Length based on the Top Attachment         
                    if (HgrCompareDoubleService.cmpdbl(channelDistance , 0)==true)
                    {
                        if (maxAngularMovement > 5 * (3.14 / 180))    // check more than 5deg
                        {
                            //Lug and Clevis Attachment
                            if (adjustable == "Yes" || (adjustable == "Auto" && (length - clamp_to - eyeNutTo - clevisTo - lugTo >= 0.8382)))
                            {
                                midRodLength = length - clamp_to - eyeNutTo - clevisTo - lugTo - 0.2286;
                                bottomRodLength = 0.1524;    //6in
                                turnbuckleFlag = "1";
                            }
                            if (adjustable == "No" || (adjustable == "Auto" && (length - clamp_to - eyeNutTo - clevisTo - lugTo < 0.8382)))
                                midRodLength = length - clamp_to - eyeNutTo - clevisTo - lugTo;


                            if (length - clamp_to - eyeNutTo - clevisTo - lugTo < 0.1524)
                                throw new CmnException("Rod lengths are too short. Use a Frame.");
                        }
                        else
                        {
                            //Welded Beam Attachment with Eyenut
                            if (adjustable == "Yes" || (adjustable == "Auto" && (length - clamp_to - eyeNutTo * 2.0 - beamAttachmentTo >= 0.8382)))
                            {
                                midRodLength = length - clamp_to - eyeNutTo * 2.0 - beamAttachmentTo - 0.2286;
                                bottomRodLength = 0.1524;    // 6in
                                turnbuckleFlag = "1";
                            }
                            if (adjustable == "No" || (adjustable == "Auto" && (length - clamp_to - eyeNutTo * 2.0 - beamAttachmentTo < 0.8382)))
                                midRodLength = length - clamp_to - eyeNutTo * 2.0 - beamAttachmentTo;

                            if (length - clamp_to - eyeNutTo * 2.0 - beamAttachmentTo < 0.1524)
                                throw new CmnException("Rod lengths are too short. Use a Frame.");
                        }
                    }
                    else
                    {                        //Washer Plate Attachment
                        if (adjustable == "Yes")
                        {
                            midRodLength = length - clamp_to - eyeNutTo + rodDiameter * 3.0 - 0.2286;
                            bottomRodLength = 0.1524;    //6in
                            turnbuckleFlag = "1";
                        }
                        else
                            midRodLength = length - clamp_to - eyeNutTo + rodDiameter * 3;

                        if (length - clamp_to - eyeNutTo < 0.1524)
                            throw new CmnException("Rod lengths are too short. Use a Frame.");
                    }
                    if (midRodLength > CONST_1)  //'144in
                    {
                        trigger = "2";
                        noOfRods = Convert.ToInt32(midRodLength / CONST_1) + 1;
                        shortRodLength = midRodLength - (noOfRods - 1) * CONST_1;
                        noOfCouplers = noOfRods - 1;
                    }

                    foreach (BusinessObject part in fig140NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            midRodWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG140N", "WEIGHT_PER_LENGTH")).PropValue;
                            break;
                        }
                    }
                    PartClass lrParts_fig253NPartClass = (PartClass)catalogBaseHelper.GetPartClass("LRParts_FIG253N");
                    fig253NParts = lrParts_fig253NPartClass.Parts;
                    foreach (BusinessObject part in fig253NParts)
                    {
                        if ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG253N", "ROD_DIA")).PropValue >= rodDiameterOverride - 0.001 && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG253N", "ROD_DIA")).PropValue <= rodDiameterOverride + 0.001)
                        {
                            botRodWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrLRParts_FIG253N", "WEIGHT_PER_LENGTH")).PropValue;
                            break;
                        }
                    }
                    if ((rodDiameter >= (0.009525 - 0.001) && rodDiameter <= (0.009525 + 0.001))) inchRodDiameter = 0.375;
                    if ((rodDiameter >= (0.0127 - 0.001) && rodDiameter <= (0.0127 + 0.001))) inchRodDiameter = 0.5;
                    if ((rodDiameter >= (0.015875 - 0.001) && rodDiameter <= (0.015875 + 0.001))) inchRodDiameter = 0.625;
                    if ((rodDiameter >= (0.01905 - 0.001) && rodDiameter <= (0.01905 + 0.001))) inchRodDiameter = 0.75;
                    if ((rodDiameter >= (0.022225 - 0.001) && rodDiameter <= (0.022225 + 0.001))) inchRodDiameter = 0.875;
                    if ((rodDiameter >= (0.0254 - 0.001) && rodDiameter <= (0.0254 + 0.001))) inchRodDiameter = 1;
                    if ((rodDiameter >= (0.028575 - 0.001) && rodDiameter <= (0.028575 + 0.001))) inchRodDiameter = 1.125;
                    if ((rodDiameter >= (0.03175 - 0.001) && rodDiameter <= (0.03175 + 0.001))) inchRodDiameter = 1.25;
                    if ((rodDiameter >= (0.0381 - 0.001) && rodDiameter <= (0.0381 + 0.001))) inchRodDiameter = 1.5;
                    if ((rodDiameter >= (0.04445 - 0.001) && rodDiameter <= (0.04445 + 0.001))) inchRodDiameter = 1.75;
                    if ((rodDiameter >= (0.0508 - 0.001) && rodDiameter <= (0.0508 + 0.001))) inchRodDiameter = 2;
                    if ((rodDiameter >= (0.05715 - 0.001) && rodDiameter <= (0.05715 + 0.001))) inchRodDiameter = 2.25;
                    if ((rodDiameter >= (0.0635 - 0.001) && rodDiameter <= (0.0635 + 0.001))) inchRodDiameter = 2.5;

                    parts.Add(new PartInfo(CLAMP_TYPE, clampType));
                     //**************************************************************
                    parts.Add(new PartInfo(BOT_EYENUT_1, "LRParts_FIG290N" + "_" + inchRodDiameter));
                    if (turnbuckleFlag == "1" || trigger == "2")
                    {
                        hexnut_counter++;
                        parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                    }
                    //************************************************************************************************************************************************************************
                    //Rods and Rod Connectors
                    if (turnbuckleFlag == "1")
                    {
                        parts.Add(new PartInfo(BOT_ROD_1, "LRParts_FIG253N" + "_" + inchRodDiameter));
                        parts.Add(new PartInfo(TURNBUCKLE_1, "LRParts_FIG230N" + "_" + inchRodDiameter));
                        hexnut_counter++;
                        parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                        if (trigger == "2")
                        {
                            midrod_counter++;
                            parts.Add(new PartInfo(MID_ROD + "_" + midrod_counter, "LRParts_FIG140N" + "_" + inchRodDiameter));
                            for (int Counter = 1; Counter <= noOfRods - 1; Counter++)
                            {
                                midrod_counter++; coupling_counter++;
                                parts.Add(new PartInfo(MID_ROD + "_" + midrod_counter, "LRParts_FIG140N" + "_" + inchRodDiameter));
                                parts.Add(new PartInfo(COUPLING + "_" + coupling_counter, "LRParts_FIG135N" + "_" + inchRodDiameter));
                                if (pipeDiameter > CONST_PIPEDIAMETER)
                                {
                                    hexnut_counter++;
                                    parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                                    hexnut_counter++;
                                    parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                                }
                                else
                                {
                                    hexnut_counter++;
                                    parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                                }
                            }
                        }
                        else
                        {
                            midrod_counter++;
                            parts.Add(new PartInfo(MID_ROD + "_" + midrod_counter, "LRParts_FIG140N" + "_" + inchRodDiameter));
                        }
                    }
                    else
                    {
                        if (trigger == "2")
                        {
                            midrod_counter++;
                            parts.Add(new PartInfo(MID_ROD + "_" + midrod_counter, "LRParts_FIG140N" + "_" + inchRodDiameter));
                            for (int Counter = 1; Counter <= noOfRods - 1; Counter++)
                            {
                                midrod_counter++; coupling_counter++;
                                parts.Add(new PartInfo(MID_ROD + "_" + midrod_counter, "LRParts_FIG140N" + "_" + inchRodDiameter));
                                parts.Add(new PartInfo(COUPLING + "_" + coupling_counter, "LRParts_FIG135N" + "_" + inchRodDiameter));
                                if (pipeDiameter > CONST_PIPEDIAMETER)
                                {
                                    hexnut_counter++;
                                    parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                                    hexnut_counter++;
                                    parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                                }
                                else
                                {
                                    hexnut_counter++;
                                    parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                                }
                            }
                        }
                        else
                        {
                            midrod_counter++;
                            parts.Add(new PartInfo(MID_ROD + "_" + midrod_counter, "LRParts_FIG140N" + "_" + inchRodDiameter));
                        }
                    }
                    //************************************************************************************************************************************************************************
                    //Top Attachment
                    if (HgrCompareDoubleService.cmpdbl(channelDistance, 0) == true)
                    {
                        if (pipeDiameter > CONST_PIPEDIAMETER && turnbuckleFlag == "1")
                        {
                            hexnut_counter++;
                            parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                        }
                        if (turnbuckleFlag != "1")
                        {
                            hexnut_counter++;
                            parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                        }
                        if (maxAngularMovement > 5 * (3.14 / 180))
                        {
                            parts.Add(new PartInfo(CLEVIS_1, "LRParts_FIG299N" + "_" + inchRodDiameter));
                            parts.Add(new PartInfo(LUG_1, "LRParts_FIG55NS" + "_" + inchRodDiameter));
                            supportingPartKey = LUG_1;
                        }
                        else
                        {
                            parts.Add(new PartInfo(EYE_NUT_1, "LRParts_FIG290N" + "_" + inchRodDiameter));
                            parts.Add(new PartInfo(WBA_1, "LRParts_FIG66N" + "_" + inchRodDiameter));
                            supportingPartKey = WBA_1;
                        }
                    }
                    else
                    {
                        //Washer Plate Attachment
                        channelDistance = 1;
                        if (channelDistance < 2 * rodDiameter)
                            throw new CmnException("Distance between Structural Channels is too small.");
                        hexnut_counter++;
                        parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                        hexnut_counter++;
                        parts.Add(new PartInfo(HEX_NUT + "_" + hexnut_counter, "LRParts_HEX_NUT" + "_" + inchRodDiameter));
                        parts.Add(new PartInfo(WPLATE_1, "LRParts_FIG60N" + "_" + inchRodDiameter));
                        supportingPartKey = WPLATE_1;
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
                hexnut_counter = 0; midrod_counter = 0; coupling_counter = 0;

                //====== ======
                //Create Joints
                //====== ======
                //Joint from Pipe Clamp to Pipe
                JointHelper.CreateRigidJoint(CLAMP_TYPE, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                //****************************************************************************************
                //For Dimension Note
                ControlPoint controlPoint;
                Note noteStart,noteEnd;
                noteStart = CreateNote("L_Start", CLAMP_TYPE, "Route", new Position(0, 0, 0), " ", true, 2, 1, out controlPoint);
                noteEnd = CreateNote("L_End", CLAMP_TYPE, "Route", new Position(0, 0, length), " ", true, 2, 1, out controlPoint);
                //*****************************************
                // L Note
                //*****************************************
                noteStart.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem fabrication = noteStart.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                noteStart.SetPropertyValue(fabrication, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                noteStart.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                noteEnd.SetPropertyValue("", "IJGeneralNote", "Text");
                fabrication = noteEnd.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                noteEnd.SetPropertyValue(fabrication, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                noteEnd.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                //Add a Revolute Joint between Eye Nut and Pipe Clamp
                JointHelper.CreateRigidJoint(BOT_EYENUT_1, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, (clamp_to + eyeNutTo), 0, 0);
                if (turnbuckleFlag == "1" || trigger == "2")
                {
                    hexnut_counter++;
                    JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + rodDiameter * 3.5, 0, 0);
                }
                //************************************************************************************************************************************************************************
                //Rods and Rod Connectors
                if (turnbuckleFlag == "1")
                {
                    JointHelper.CreateRigidJoint(BOT_ROD_1, "ExThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, (clamp_to + eyeNutTo), 0, 0);
                    componentDictionary[BOT_ROD_1].SetPropertyValue(bottomRodLength, "IJUAHgrOccLength", "Length");
                    JointHelper.CreateRigidJoint(TURNBUCKLE_1, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo, 0, 0);
                    hexnut_counter++;
                    JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo * 2.0, 0, 0);
                    if (trigger == "2")
                    {
                        midrod_counter++;
                        JointHelper.CreateRigidJoint(MID_ROD + "_" + midrod_counter, "TopExThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo, 0, 0);
                        componentDictionary[MID_ROD + "_" + midrod_counter].SetPropertyValue(shortRodLength, "IJUAHgrOccLength", "Length");
                        for (int Counter = 1; Counter <= noOfRods - 1; Counter++)
                        {
                            midrod_counter++; coupling_counter++;
                            JointHelper.CreateRigidJoint(MID_ROD + "_" + midrod_counter, "BotExThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo + shortRodLength + 144 * (Counter - 1), 0, 0);
                            componentDictionary[MID_ROD + "_" + midrod_counter].SetPropertyValue(CONST_1, "IJUAHgrOccLength", "Length"); //144in
                            JointHelper.CreateRigidJoint(COUPLING + "_" + coupling_counter, "BotInThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo + shortRodLength + 144 * (Counter - 1), 0, 0);
                            if (pipeDiameter > CONST_PIPEDIAMETER)
                            {
                                hexnut_counter++;
                                JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo + shortRodLength - rodDiameter * 3.0 + CONST_1 * (Counter - 1), 0, 0);

                                hexnut_counter++;
                                JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo + shortRodLength + rodDiameter * 3.0 + CONST_1 * (Counter - 1), 0, 0);
                            }
                            else
                            {
                                hexnut_counter++;
                                JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo + shortRodLength + rodDiameter * 3.0 + CONST_1 * (Counter - 1), 0, 0);
                            }
                        }
                    }
                    else
                    {
                        midrod_counter++;
                        JointHelper.CreateRigidJoint(MID_ROD + "_" + midrod_counter, "TopExThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + bottomRodLength + turnbuckleTo, 0, 0);
                        ((SupportComponent)componentDictionary[MID_ROD + "_" + midrod_counter]).SetPropertyValue(midRodLength, "IJUAHgrOccLength", "Length");
                    }
                }
                else
                {
                    if (trigger == "2")
                    {
                        midrod_counter++;
                        JointHelper.CreateRigidJoint(MID_ROD + "_" + midrod_counter, "BotExThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo, 0, 0);
                        ((SupportComponent)componentDictionary[MID_ROD + "_" + midrod_counter]).SetPropertyValue(shortRodLength, "IJUAHgrOccLength", "Length");

                        for (int Counter = 1; Counter <= noOfRods - 1; Counter++)
                        {
                            midrod_counter++; coupling_counter++;
                            JointHelper.CreateRigidJoint(MID_ROD + "_" + midrod_counter, "BotExThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + shortRodLength + CONST_1 * (Counter - 1), 0, 0);
                            ((SupportComponent)componentDictionary[MID_ROD + "_" + midrod_counter]).SetPropertyValue(CONST_1, "IJUAHgrOccLength", "Length"); //144in
                            JointHelper.CreateRigidJoint(COUPLING + "_" + coupling_counter, "BotInThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + shortRodLength + CONST_1 * (Counter - 1), 0, 0);
                            if (pipeDiameter > CONST_PIPEDIAMETER)
                            {
                                hexnut_counter++;
                                JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + shortRodLength - rodDiameter * 3.0 + CONST_1 * (Counter - 1), 0, 0);

                                hexnut_counter++;
                                JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + shortRodLength + rodDiameter * 3.0 + CONST_1 * (Counter - 1), 0, 0);
                            }
                            else
                            {
                                hexnut_counter++;
                                JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo + shortRodLength + rodDiameter * 3.0 + CONST_1 * (Counter - 1), 0, 0);

                            }
                        }
                    }
                    else
                    {
                        midrod_counter++;
                        JointHelper.CreateRigidJoint(MID_ROD + "_" + midrod_counter, "BotExThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, clamp_to + eyeNutTo, 0, 0);
                        ((SupportComponent)componentDictionary[MID_ROD + "_" + midrod_counter]).SetPropertyValue(midRodLength, "IJUAHgrOccLength", "Length"); //144in


                    }
                }
                //************************************************************************************************************************************************************************
                //Top Attachment
                if (HgrCompareDoubleService.cmpdbl(channelDistance, 0) == true)
                {
                    if (maxAngularMovement > 5 * (3.14 / 180))
                    {
                        if (pipeDiameter > CONST_PIPEDIAMETER && turnbuckleFlag == "1")
                        {
                            hexnut_counter++;
                            JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length - clevisTo - lugTo - rodDiameter * 3.0, 0, 0);
                        }
                        if (turnbuckleFlag != "1")
                        {
                            hexnut_counter++;
                            JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length - clevisTo - lugTo - rodDiameter * 3.0, 0, 0);
                        }
                        JointHelper.CreateRigidJoint(CLEVIS_1, "Pin", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length - lugTo, 0, 0);
                        JointHelper.CreateRigidJoint(LUG_1, "Structure", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, length, 0, 0);
                    }
                    else
                    {
                        //Welded Beam Attachment with Eyenut
                        if (pipeDiameter > CONST_PIPEDIAMETER && turnbuckleFlag == "1")
                        {
                            hexnut_counter++;
                            JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length - beamAttachmentTo - eyeNutTo - rodDiameter * 3.0, 0, 0);
                        }
                        if (turnbuckleFlag != "1")
                        {
                            hexnut_counter++;
                            JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length - beamAttachmentTo - eyeNutTo - rodDiameter * 3.0, 0, 0);
                        }

                        JointHelper.CreateRigidJoint(EYE_NUT_1, "Eye", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length - beamAttachmentTo, 0, 0);
                        JointHelper.CreateRigidJoint(WBA_1, "Structure", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, length, 0, 0);
                    }
                }
                else
                {
                    //Washer Plate Attachment
                    channelDistance = 1;
                    if (channelDistance < 2 * rodDiameter)
                        throw new CmnException("Distance between Structural Channels is too small.");
                    hexnut_counter++;
                    JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length + rodDiameter * 2.0, 0, 0);
                    hexnut_counter++;
                    JointHelper.CreateRigidJoint(HEX_NUT + "_" + hexnut_counter, "InThdRH", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, length + rodDiameter * 2, 0, 0);
                    JointHelper.CreateRigidJoint(WPLATE_1, "Structure", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, length, 0, 0);
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
                    routeConnections.Add(new ConnectionInfo(CLAMP_TYPE, 1));
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
                    structConnections.Add(new ConnectionInfo(supportingPartKey, 1));

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

