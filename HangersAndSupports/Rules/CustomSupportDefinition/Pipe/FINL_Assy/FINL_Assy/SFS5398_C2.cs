//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5398_C2.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5398_C2
//   Author       :  Vijay
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     BS   CR-CP-224492- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5398_C2 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        string size4;
        PropertyValueCodelist size4CodeList;
        double shoeH, X1, X2, steelOverlap, horizontalOverLap, normalPipeDiameter;

        private const string HORSTEEL = "HorSteel_5398_C2";
        private const string VERSTEEL = "VerSteel_5398_C2";
        private const double SteelDensityKGPerM = 7900;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    size4CodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize4", "Size4");
                    size4 = size4CodeList.PropertyInfo.CodeListInfo.GetCodelistItem(size4CodeList.PropValue).DisplayName;
                    steelOverlap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLSteelOverlap", "SteelOverlap")).PropValue;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    X1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLX1", "X1")).PropValue;
                    X2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLX2", "X2")).PropValue;
                    horizontalOverLap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLHorOverlap", "HorOverlap")).PropValue;

                    parts.Add(new PartInfo(HORSTEEL, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(VERSTEEL, "Util_Fixed_Box_Metric_1"));
                    return parts;       //Get the collection of parts
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                //To get Pipe Nom Dia
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                if (pipeInfo.NominalDiameter.Units != "mm")
                    normalPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER) / 1000;
                else
                    normalPipeDiameter = pipeInfo.NominalDiameter.Size / 1000;

                NominalDiameter minNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                NominalDiameter maxNominalDiameter = new NominalDiameter();
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5398_C2.cs", 113);
                    return;
                }

                //Getting Dimension Information
                //Interface name: IJUAFINLSrv_SFS5398_C2_Dim
                PropertyValue TSSize1, TSSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size4", size4);

                GenericHelper.GetDataByRule("FINLSrv_SFS5398_C2_Dim", "IJUAFINLSrv_SFS5398_C2_Dim", "TS_Size1", parameter, out TSSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5398_C2_Dim", "IJUAFINLSrv_SFS5398_C2_Dim", "TS_Size2", parameter, out TSSize2);
                double maxSpan = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_C2_Dim", "IJUAFINLSrv_SFS5398_C2_Dim", "Max_Span", "IJUAFINLSrv_SFS5398_C2_Dim", "Size4", size4);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_C2_Dim", "IJUAFINLSrv_SFS5398_C2_Dim", "Max_Len", "IJUAFINLSrv_SFS5398_C2_Dim", "Size4", size4);
                double MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_C2_Dim", "IJUAFINLSrv_SFS5398_C2_Dim", "MA", "IJUAFINLSrv_SFS5398_C2_Dim", "Size4", size4);
                double MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_C2_Dim", "IJUAFINLSrv_SFS5398_C2_Dim", "MP", "IJUAFINLSrv_SFS5398_C2_Dim", "Size4", size4);
                double FMax = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_C2_Dim", "IJUAFINLSrv_SFS5398_C2_Dim", "Fmax", "IJUAFINLSrv_SFS5398_C2_Dim", "Size4", size4);

                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(FMax, "IJUAFINLFmax", "Fmax");

                double horizontalLength, verticalLength, widthHorTS, widthVerTS, depthHorTS, depthVerTS, existSteelD = 0;

                //add two C shape guides  -UPN100
                //Get steel Data
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                string sectionStandardC = "Euro";
                CrossSection crossSection1 = catalogStructHelper.GetCrossSection(sectionStandardC, Convert.ToString(TSSize1));
                widthHorTS = crossSection1.Width;
                depthHorTS = crossSection1.Depth;

                CrossSection crossSection2 = catalogStructHelper.GetCrossSection(sectionStandardC, Convert.ToString(TSSize2));
                widthVerTS = crossSection2.Width;
                depthVerTS = crossSection2.Depth;

                horizontalLength = X1 + X2 + horizontalOverLap;
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


                if (supportingType == "Slab")
                    existSteelD = 0;
                else if (supportingType == "Steel")
                    existSteelD = SupportingHelper.SupportingObjectInfo(1).Depth;
             

                double Angleoffset, theta;

                theta = Math.Abs((Math.Atan(1) * 4.0) / 2 - RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z));
                Angleoffset = Math.Tan(theta) * widthHorTS;

                double verticalDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                verticalLength = verticalDistance + pipeInfo.OutsideDiameter / 2 + shoeH + depthHorTS + steelOverlap - existSteelD / 2 + Angleoffset / 2;

                //Checking for max H
                if (verticalDistance + shoeH + pipeInfo.OutsideDiameter / 2 + depthHorTS > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WANRING: " + "Maximum allowable length of " + Convert.ToString(maxLength * 1000) + "mm exceeded.", "", "SFS5398_C2.cs", 167);

                //If the span length entered is greater than allowed
                if (X1 + X2 > maxSpan)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable span of " + Convert.ToString(maxLength * 1000) + "mm exceeded.", "", "SFS5398_C2.cs", 172);
                    return;
                }

                if (steelOverlap < (100.0 / 1000.0))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Minimum Vertical Overlap is " + Convert.ToString(100) + " mm", "", "SFS5398_C2.cs", 178);
                    return;
                }

                if (horizontalOverLap < (100.0 / 1000.0))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Minimum Horizontal Overlap is " + Convert.ToString(100) + " mm", "", "SFS5398_C2.cs", 184);
                    return;
                }

                if (pipeInfo.OutsideDiameter / 2 > X1 || pipeInfo.OutsideDiameter / 2 > X2)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "X1 and X2 must exceed Pipe outside radius.", "", "SFS5398_C2.cs", 190);
                    return;
                }

                double verticalSteelWeight = (widthVerTS * depthVerTS * verticalLength - ((widthVerTS - 2.0 * 5.0 / 1000) * (depthVerTS - 2.0 * 5.0 / 1000) * verticalLength)) * SteelDensityKGPerM;
                double horizontalSteelWeight;

                if (size4CodeList.PropValue == 3)
                    horizontalSteelWeight = (widthHorTS * depthHorTS * horizontalLength - ((widthHorTS - 2.0 * 5.0 / 1000) * (depthHorTS - 2.0 * 5.0 / 1000) * horizontalLength)) * SteelDensityKGPerM;
                else
                    horizontalSteelWeight = (widthHorTS * depthHorTS * horizontalLength - ((widthHorTS - 2.0 * 4.0 / 1000) * (depthHorTS - 2.0 * 4.0 / 1000) * horizontalLength)) * SteelDensityKGPerM;

                componentDictionary[HORSTEEL].SetPropertyValue(widthHorTS, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[HORSTEEL].SetPropertyValue(depthHorTS, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[HORSTEEL].SetPropertyValue(horizontalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[HORSTEEL].SetPropertyValue(TSSize1 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0), "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[HORSTEEL].SetPropertyValue(horizontalSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                componentDictionary[VERSTEEL].SetPropertyValue(widthVerTS, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[VERSTEEL].SetPropertyValue(depthVerTS, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[VERSTEEL].SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[VERSTEEL].SetPropertyValue(TSSize2 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0), "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[VERSTEEL].SetPropertyValue(verticalSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                //Create graphics
                if (supportingType == "Steel")
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, verticalDistance + pipeInfo.OutsideDiameter / 2 + shoeH + depthHorTS / 2 + Angleoffset / 2, 0, -X2);
                    else
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WANRING: " + "Support can only be placed on steel By-Structure.", "", "SFS5398_C2.cs", 220);
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, (verticalDistance + pipeInfo.OutsideDiameter / 2 + shoeH + depthHorTS / 2 + Angleoffset / 2), X2, 0);
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY, (verticalDistance + pipeInfo.OutsideDiameter / 2 + shoeH + depthHorTS / 2 + Angleoffset / 2), -X2, 0);
                    }
                    else
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WANRING: " + "Support can only be placed on Slab By-Point.", "", "SFS5398_C2.cs", 232);
                }

                JointHelper.CreateRigidJoint(HORSTEEL, "StartOther", VERSTEEL, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, -depthVerTS / 2, 0, -depthHorTS / 2);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Finl_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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

                    routeConnections.Add(new ConnectionInfo(HORSTEEL, 1));      //partindex, routeindex

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(VERSTEEL, 1));     //partindex, structindex

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

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string BOMString = "";
            try
            {
                //To get Pipe Nom Dia
                int sizeBOM = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLSize4", "Size4")).PropValue;

                BOMString = "Gate support C2 -" + Convert.ToString(sizeBOM) + " SFS 5398";

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        #endregion
    }
}
