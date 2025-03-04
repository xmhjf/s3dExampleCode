//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   StrutAssemblyServices.cs
//  
//   Author       :  PVK
//   Creation Date:  04-11-2014
//   Description:

//   Change History:
//   dd.mmm.yyyy     who        change description
//   -----------     ---        ------------------
//   04-Nov-2014     PVK        CR-CP-245790 Modify the exsisting .Net Strut_Assy to new URS Strut supports
//   22-Jan-2015     PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report 
//   28-04-2015      PVK	 	Resolve Coverity issues found in April
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Route.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    public static class StrutAssemblyServices
    {
        /// <summary>
        /// Data Type For Weld Parts
        /// </summary>
        public struct WeldData
        {
            public string partKey;
            public string partNumber;
            public string partRule;
            public string connection;
            public int location;
            public double offsetXValue;
            public double offsetYValue;
            public double offsetZValue;
        }

        /// <summary>
        /// This method will be called to check the support with the given family and type attributes
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="family"> the value of family attribute from support</param>
        /// <param name="type">the value of type attribute from support</param>
        /// <returns></returns>
        /// <code>
        /// CheckSupportWithFamilyAndType(this, family, type)
        /// </code>
        public static void CheckSupportWithFamilyAndType(CustomSupportDefinition customSupportDefinition, string family, string type)
        {
            try
            {
                RouteFeature routeFeature = customSupportDefinition.SupportHelper.SupportedObjects[0] as RouteFeature;
                ReadOnlyCollection<RoutePart> collection = null;
                if(routeFeature!=null)
                    collection = routeFeature.Parts;
                RoutePart genPart = null;
                if (collection != null)
                {
                    genPart = collection[0] as RoutePart;
                    genPart = collection[0];
                }

                if (family == "" || type == "")
                    return;

                double familyNDPFrom = 0, familyNDPTo = 0;
                string geStdWeight = string.Empty, lStdWeight = string.Empty, carbonSteel = string.Empty, alloySteel = string.Empty, stainlessSteel = string.Empty, LE3 = string.Empty, G3 = string.Empty, PP = string.Empty, FP = string.Empty, HC = string.Empty, CC = string.Empty, freezeProt = string.Empty, procHeat = string.Empty, controTrace = string.Empty, t300 = string.Empty, t650 = string.Empty, gt650 = string.Empty, t750 = string.Empty, gt750 = string.Empty;
                // Get configuration variable from excel sheet.
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass ursWGFamilyAux = (PartClass)catalogBaseHelper.GetPartClass("hsURS_WG_Family");
                ReadOnlyCollection<BusinessObject> ursWGFamilyItems = ursWGFamilyAux.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                Boolean isWgClass = false;
                if (ursWGFamilyItems != null)
                {
                    foreach (BusinessObject part1 in ursWGFamilyItems)
                    {
                        if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Family")).PropValue == family) && ((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Type")).PropValue == type))
                        {
                            familyNDPFrom = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaFrom")).PropValue;
                            familyNDPTo = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaTo")).PropValue;
                            geStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GeStdWeight")).PropValue;
                            lStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LStdWeight")).PropValue;
                            carbonSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CS")).PropValue;
                            alloySteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "AS")).PropValue;
                            stainlessSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "SS")).PropValue;
                            LE3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LE3")).PropValue;
                            G3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "G3")).PropValue;
                            PP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "PP")).PropValue;
                            FP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FP")).PropValue;
                            HC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "HC")).PropValue;
                            CC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CC")).PropValue;
                            freezeProt = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FreezeProt")).PropValue;
                            procHeat = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ProcHeat")).PropValue;
                            controTrace = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ControTrace")).PropValue;
                            t300 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T300")).PropValue;
                            t650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T650")).PropValue;
                            gt650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT650")).PropValue;
                            t750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T750")).PropValue;
                            gt750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT750")).PropValue;
                            isWgClass = true;
                            break;
                        }
                    }
                }
                if (isWgClass == false)
                {
                    ursWGFamilyAux = (PartClass)catalogBaseHelper.GetPartClass("hsURSAssy_WG_NFamily");
                    ursWGFamilyItems = ursWGFamilyAux.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    if (ursWGFamilyItems != null)
                    {
                        foreach (BusinessObject part1 in ursWGFamilyItems)
                        {
                            if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Family")).PropValue == family) && ((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Type")).PropValue == type))
                            {
                                familyNDPFrom = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaFrom")).PropValue;
                                familyNDPTo = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaTo")).PropValue;
                                geStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GeStdWeight")).PropValue;
                                lStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LStdWeight")).PropValue;
                                carbonSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CS")).PropValue;
                                alloySteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "AS")).PropValue;
                                stainlessSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "SS")).PropValue;
                                LE3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LE3")).PropValue;
                                G3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "G3")).PropValue;
                                PP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "PP")).PropValue;
                                FP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FP")).PropValue;
                                HC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "HC")).PropValue;
                                CC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CC")).PropValue;
                                freezeProt = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FreezeProt")).PropValue;
                                procHeat = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ProcHeat")).PropValue;
                                controTrace = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ControTrace")).PropValue;
                                t300 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T300")).PropValue;
                                t650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T650")).PropValue;
                                gt650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT650")).PropValue;
                                t750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T750")).PropValue;
                                gt750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT750")).PropValue;
                                break;
                            }
                        }
                    }
                }
                // Get all pipes on the support.
                // *********************************************************************************************
                // Filter based on Pipe Size, command type, discipline type
                // *********************************************************************************************
                PipeObjectInfo pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(1);
                double primaryPipeSize = pipeInfo.NominalDiameter.Size;
                string unit = pipeInfo.NominalDiameter.Units;

                if (unit == "mm")
                    primaryPipeSize = Math.Round(MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, primaryPipeSize, UnitName.NPD_MILLIMETER, UnitName.NPD_INCH), 4);
                // check valid pipe size
                if (HgrCompareDoubleService.cmpdbl(familyNDPFrom, 0) == false && HgrCompareDoubleService.cmpdbl(familyNDPTo, 0) == false)
                {
                    if (familyNDPFrom > primaryPipeSize || familyNDPTo < primaryPipeSize)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Pipe size not valid.", "", "URSAssemblyServices.cs", 557);
                }

                // Setting up to get pipe information that we will use to determine what supports can be used.
                // Get the insulation thickness of the pipe that we are placing on.
                double inslatThickness = pipeInfo.InsulationThickness;
                double compareThickness = Math.Round(MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, inslatThickness, UnitName.DISTANCE_METER, UnitName.DISTANCE_INCH), 0);
                string insulCompare = string.Empty;
                if (compareThickness <= 3)
                    insulCompare = "LE3";
                else
                    insulCompare = "G3";

                int inslatCodeTemp = 0, inslatGroupCode = 0;
                string inslatCode = string.Empty;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (HgrCompareDoubleService.cmpdbl(inslatThickness ,0)==false)
                {
                    // Get the index number for the inslation purpose
                    inslatCodeTemp = pipeInfo.InsulationPurpose;
                    // Convert the index number into the string representing the purpose
                    inslatGroupCode = metadataManager.GetCodelistInfo("InsulationPurpose", "REFDAT").GetCodelistItem(inslatCodeTemp).ParentValue;
                    inslatCode = metadataManager.GetCodelistInfo("InsulationType", "REFDAT").GetCodelistItem(inslatCodeTemp).ShortDisplayName.Trim();

                    // Format the insulation code to match what we have in the excel sheet.
                    if (inslatCode == "Heat conservation")
                        inslatCode = "HC";
                    else if (inslatCode == "Cold conservation")
                        inslatCode = "CC";
                    else if (inslatCode == "Fire proofing")
                        inslatCode = "FP";
                    else if (inslatCode == "Safety")
                        inslatCode = "PP";
                }

                string lineMaterial = customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).MaterialCategory;

                if (lineMaterial == "Carbon Steels")
                    lineMaterial = "CarbonSteel";
                else
                {
                    if (lineMaterial == "Stainless Steels")
                        lineMaterial = "StainlessSteel";
                    else
                        lineMaterial = "AlloySteel";
                }

                // Get the pipe temperature
                double pipeTemp = pipeInfo.MaxDesignTemperature;
                ReadOnlyCollection<IPort> portcollection;
                IPipePort pipeport;
                int scheduleValue = 0;
                string schedule = string.Empty;
                double pipeWallThickness = 0;
                if (genPart.GetType().FullName.Equals("Ingr.SP3D.Route.Middle.PipeComponent"))
                {
                    Ingr.SP3D.Route.Middle.PipeComponent pipecomponent = (Ingr.SP3D.Route.Middle.PipeComponent)genPart;
                    portcollection = pipecomponent.GetConnectedPorts(PortType.Piping);
                    pipeport = (IPipePort)portcollection[1];
                    scheduleValue = pipeport.ScheduleThickness;
                    schedule = metadataManager.GetCodelistInfo("ScheduleThickness", "REFDAT").GetCodelistItem(scheduleValue).ShortDisplayName.Trim();
                    pipeWallThickness = pipeport.WallThicknessOrGrooveSetback;
                }
                else
                {
                    Ingr.SP3D.Route.Middle.PipeStockPart pipeStockPart = (Ingr.SP3D.Route.Middle.PipeStockPart)genPart;
                    portcollection = pipeStockPart.GetConnectablePorts(PortType.Piping);
                    pipeport = (IPipePort)portcollection[1];
                    scheduleValue = pipeport.ScheduleThickness;
                    schedule = metadataManager.GetCodelistInfo("ScheduleThickness", "REFDAT").GetCodelistItem(scheduleValue).ShortDisplayName.Trim();
                    pipeWallThickness = pipeport.WallThicknessOrGrooveSetback;
                }


                string schedualCompare = string.Empty;
                double wallThickness = 0;
                if (schedule.ToUpper() == "UNDEFINED")
                    schedualCompare = "Unknown";
                else if (schedule.ToUpper() != "S-STD")
                {
                    catalogBaseHelper = new CatalogBaseHelper();
                    ursWGFamilyAux = (PartClass)catalogBaseHelper.GetPartClass("REFDATPlainPipeEndData");
                    ursWGFamilyItems = ursWGFamilyAux.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    if (ursWGFamilyItems != null)
                    {
                        foreach (BusinessObject part1 in ursWGFamilyItems)
                        {
                            if ((HgrCompareDoubleService.cmpdbl((double)((PropertyValueDouble)part1.GetPropertyValue("IJPlainPipeEndData", "NominalPipingDiameter")).PropValue , primaryPipeSize)==true) && ((string)((PropertyValueString)part1.GetPropertyValue("IJPlainPipeEndData", "Schedule")).PropValue == "100"))
                            {
                                wallThickness = (double)((PropertyValueDouble)part1.GetPropertyValue("IJPlainPipeEndData", "WallThickness")).PropValue;
                                if (pipeWallThickness < wallThickness)
                                    schedualCompare = "LStdWeight";
                                else
                                    schedualCompare = "GeStdWeight";
                            }
                        }
                    }
                }
                else
                    schedualCompare = "GeStdWeight";

                string tempCode = string.Empty;
                // these will compare the tempature of the pipe with the allowed tempature.  We will be using the Kelvin units of mesure to compare with.
                if ((string.IsNullOrEmpty(t650) || t650.ToUpper() != "YES") && (string.IsNullOrEmpty(gt650) || gt650.ToUpper() != "YES"))
                {
                    if (pipeTemp <= 422.0389)
                        tempCode = "T300";
                    else if (pipeTemp > 422.0389 && pipeTemp <= 672.0389)
                        tempCode = "T750";
                    else
                        tempCode = "GT750";
                }
                else
                {
                    if (pipeTemp <= 422.0389)
                        tempCode = "T300";
                    else if (pipeTemp > 422.0389 && pipeTemp <= 616.4833)
                        tempCode = "T650";
                    else
                        tempCode = "GT650";
                }

                // Added a check for Unknown because the schedule is 'Undefined' for Tee's - RCM
                if (schedualCompare == "Unknown")
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Wall Thickness: Unable to determine Wall Thickness.", "", "URSAssemblyServices.cs", 659);
                else if (schedualCompare == "LStdWeight")
                {
                    if (string.IsNullOrEmpty(lStdWeight) || lStdWeight.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(lStdWeight) && lStdWeight.Trim() != "")
                        {
                            if (lStdWeight.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Wall Thickness.", "", "URSAssemblyServices.cs", 667);
                            else if (lStdWeight.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Wall Thickness.", "", "URSAssemblyServices.cs", 669);
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(geStdWeight) || geStdWeight.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(geStdWeight) && geStdWeight.Trim() != "")
                        {
                            if (geStdWeight.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Wall Thickness.", "", "URSAssemblyServices.cs", 680);
                            else if (geStdWeight.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Wall Thickness.", "", "URSAssemblyServices.cs", 682);
                        }
                    }
                }

                if (lineMaterial == "CarbonSteel")
                {
                    if (string.IsNullOrEmpty(carbonSteel) || carbonSteel.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(carbonSteel) && carbonSteel.Trim() != "")
                        {
                            if (carbonSteel.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Line Material.", "", "URSAssemblyServices.cs", 694);
                            else if (carbonSteel.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Line Material.", "", "URSAssemblyServices.cs", 696);
                        }
                    }
                }
                else if (lineMaterial == "StainlessSteel")
                {
                    if (string.IsNullOrEmpty(stainlessSteel) || stainlessSteel.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(stainlessSteel) && stainlessSteel.Trim() != "")
                        {
                            if (stainlessSteel.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Line Material.", "", "URSAssemblyServices.cs", 707);
                            else if (stainlessSteel.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Line Material.", "", "URSAssemblyServices.cs", 709);
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(alloySteel) || alloySteel.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(alloySteel) && alloySteel.Trim() != "")
                        {
                            if (alloySteel.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Line Material.", "", "URSAssemblyServices.cs", 720);
                            else if (alloySteel.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Line Material.", "", "URSAssemblyServices.cs", 722);
                        }
                    }
                }

                if (!((!string.IsNullOrEmpty(LE3) && LE3.ToUpper() == "NO") && (!string.IsNullOrEmpty(G3) && G3.ToUpper() == "NO") && HgrCompareDoubleService.cmpdbl(inslatThickness , 0)==true))
                //{
                //}
                //else
                {
                    if (insulCompare == "LE3")
                    {
                        if (string.IsNullOrEmpty(LE3) || LE3.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(LE3) && LE3.Trim() != "")
                            {
                                if (LE3.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Thickness.", "", "URSAssemblyServices.cs", 739);
                                else if (LE3.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Thickness.", "", "URSAssemblyServices.cs", 741);
                            }
                        }
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(G3) || G3.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(G3) && G3.Trim() != "")
                            {
                                if (G3.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Thickness.", "", "URSAssemblyServices.cs", 752);
                                else if (G3.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Thickness.", "", "URSAssemblyServices.cs", 754);
                            }
                        }
                    }

                    if (inslatCode == "HC")
                    {
                        if (string.IsNullOrEmpty(HC) || HC.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(HC) && HC.Trim() != "")
                            {
                                if (HC.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 766);
                                else if (HC.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 768);

                            }
                        }
                    }
                    else if (inslatCode == "CC")
                    {
                        if (string.IsNullOrEmpty(CC) || CC.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(CC) && CC.Trim() != "")
                            {
                                if (CC.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 780);
                                else if (CC.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 782);

                            }
                        }
                    }
                    else if (inslatCode == "FP")
                    {
                        if (string.IsNullOrEmpty(FP) || FP.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(FP) && FP.Trim() != "")
                            {
                                if (FP.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 794);
                                else if (FP.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 796);

                            }
                        }
                    }
                    else if (inslatCode == "PP")
                    {
                        if (string.IsNullOrEmpty(PP) || PP.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(PP) && PP.Trim() != "")
                            {
                                if (PP.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 808);
                                else if (PP.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "URSAssemblyServices.cs", 810);

                            }
                        }
                    }
                }

                if (tempCode == "T300")
                {
                    if (string.IsNullOrEmpty(t300) || t300.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(t300) && t300.Trim() != "")
                        {
                            if (t300.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "URSAssemblyServices.cs", 824);
                            else if (t300.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "URSAssemblyServices.cs", 826);
                        }
                    }
                }
                else if (tempCode == "T650")
                {
                    if (string.IsNullOrEmpty(t650) || t650.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(t650) && t650.Trim() != "")
                        {
                            if (t650.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "URSAssemblyServices.cs", 837);
                            else if (t650.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "URSAssemblyServices.cs", 839);
                        }
                    }
                }
                else if (tempCode == "GT650")
                {
                    if (string.IsNullOrEmpty(gt650) || gt650.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(gt650) && gt650.Trim() != "")
                        {
                            if (gt650.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "URSAssemblyServices.cs", 850);
                            else if (gt650.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "URSAssemblyServices.cs", 852);
                        }
                    }
                }
                else if (tempCode == "T750")
                {
                    if (string.IsNullOrEmpty(t750) || t750.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(t750) && t750.Trim() != "")
                        {
                            if (t750.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "URSAssemblyServices.cs", 863);
                            else if (t750.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "URSAssemblyServices.cs", 865);
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(gt750) || gt750.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(gt750) && gt750.Trim() != "")
                        {
                            if (gt750.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "URSAssemblyServices.cs", 876);
                            else if (gt750.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "URSAssemblyServices.cs", 878);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in CheckSupportWithFamilyAndType." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// Function to add all the welds supplied in the catalog for the given Support PartNumber.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="customSupportDefinition"></param>
        /// <param name="parts">The partinfo collection</param>
        /// <param name="interfaceName">The interface on which the weld information is stored on</param>
        /// <param name="supportPartNumber">Optional. The Support PartNumber to look up the welds for. If it is not specified, the current support part number will be used</param>
        ///<code>
        /// Returns
        ///A collection of hsWeldData Types, containing all the information for each individual weld, including the index that the partocc can be accesed from the part occurence collection.
        ///AddWeldsFromCatalog(this,parts,interfacename)
        ///</code>
        public static Collection<WeldData> AddStrutWeldsFromCatalog(CustomSupportDefinition customSupportDefinition, Collection<PartInfo> parts, string interfaceName, string supportPartNumber = "")
        {
            IEnumerable<BusinessObject> weldParts = null;
            try
            {
                string offsetZRule = string.Empty;
                WeldData weld = new WeldData();
                Collection<WeldData> weldCollection = new Collection<WeldData>();

                if (supportPartNumber == "")
                    supportPartNumber = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartNumber;

                // BusinessObject partClass = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartClass;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass S3DIframeWelds = (PartClass)catalogBaseHelper.GetPartClass("hsURS_RStrut_Welds");
                weldParts = S3DIframeWelds.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                weldParts = weldParts.Where(part => (string)((PropertyValueString)part.GetPropertyValue(interfaceName, "SupportPartNumber")).PropValue == supportPartNumber);
                int i = 1;
                foreach (BusinessObject part in weldParts)
                {
                    weld.partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsRStrutWelds", "WeldPartNumber")).PropValue;
                    weld.partRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsRStrutWelds", "WeldPartRule")).PropValue;
                    weld.connection = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsRStrutWelds", "Connection")).PropValue;

                    // Add the part to the Part Collection
                    weld.partKey = "weld" + i;
                    i++;
                    parts.Add(new PartInfo(weld.partKey, weld.partNumber, weld.partRule));

                    // Add the Weld Object to the Weld Collection
                    weldCollection.Add(weld);
                }
                return weldCollection;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in AddStrutWeldsFromCatalog." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (weldParts is IDisposable)
                {
                    ((IDisposable)weldParts).Dispose(); // This line will be executed
                }
            }
        }
    }
}
