//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5398_A1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5398_A1
//   Author       :  Vijay
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     BS   CR-CP-224492- Convert HS_FINL_Assy to C# .Net 
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
    public class SFS5398_A1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        int size4Value;
        string size4;
        double shoeH, X1, X2, H, normalPipeDiameter;
        private const string HORSTEEL = "HorSteel_5398_A1";
        private const string VERSTEEL1 = "VerSteel1_5398_A1";
        private const string VERSTEEL2 = "VerSteel2_5398_A1";
        private const string PLATE1 = "Plate1_5398_A1";
        private const string PLATE2 = "Plate2_5398_A1";
        private const double SteelDensityKGPerM = 7900;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    PropertyValueCodelist size4CodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize4", "Size4");
                    size4Value = (int)size4CodeList.PropValue;
                    size4 = size4CodeList.PropertyInfo.CodeListInfo.GetCodelistItem(size4Value).DisplayName;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    X1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLX1", "X1")).PropValue;
                    X2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLX2", "X2")).PropValue;
                    H = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLHeight", "H")).PropValue;

                    parts.Add(new PartInfo(HORSTEEL, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(VERSTEEL1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(VERSTEEL2, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(PLATE1, "Util_FourHolePl_Metric_1"));        //From_utility_metric
                    parts.Add(new PartInfo(PLATE2, "Util_FourHolePl_Metric_1"));        //From_utility_metric
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
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5398_A1.cs", 118);
                    return;
                }

                //Getting Dimension Information
                //Interface name: IJUAFINLSrv_SFS5398_A1_Dim
                PropertyValue TSSize1, TSSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size4", size4);

                GenericHelper.GetDataByRule("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "TS_Size1", parameter, out TSSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "TS_Size2", parameter, out TSSize2);

                double maxSpan = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "Max_Span", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "Max_Len", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double plateT = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "Plate_T", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double plateW = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "Plate_W", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double plateD = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "Plate_D", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double C = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "C", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double inset = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "Inset", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double holeDiameter = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "D", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "MA", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "MP", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double FU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "MA", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double FQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "MP", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);
                double FMax = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_A1_Dim", "IJUAFINLSrv_SFS5398_A1_Dim", "Fmax", "IJUAFINLSrv_SFS5398_A1_Dim", "Size4", size4);

                support.SetPropertyValue(FU, "IJUAFINLFQFU", "FU");
                support.SetPropertyValue(FQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(FMax, "IJUAFINLFmax", "Fmax");

                double horizontalLength, verticalLength, widthOffsetVerTS;

                //add two C shape guides  -UPN100
                //Get steel Data
               

                double widthHorTS, flangeThickness, depthHorTS, webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(TSSize1), out widthHorTS, out flangeThickness, out webThickness, out depthHorTS);
                double widthVerTS, depthVerTS, TL;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(TSSize2), out widthVerTS, out TL, out webThickness, out depthVerTS);

                horizontalLength = X1 + X2;
                widthOffsetVerTS = -depthVerTS / 2 - plateT;
                verticalLength = H + C / 2;

                //Checking for max H
                if (H > maxSpan)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WANRING: " + "Maximum allowable length of " + Convert.ToString(maxLength * 1000) + "mm exceeded.", "", "SFS5398_A1.cs", 168);

                //If the span length entered is greater than allowed
                if (X1 + X2 > maxSpan)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable span of " + Convert.ToString(maxLength * 1000) + "mm exceeded.", "", "SFS5398_A1.cs", 173);
                    return;
                }

                if (pipeInfo.OutsideDiameter / 2 > X1 || pipeInfo.OutsideDiameter / 2 > X2)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "X1 and X2 must exceed Pipe outside radius.", "", "SFS5398_A1.cs", 179);
                    return;
                }

                double verticalSteelWeight = (widthVerTS * depthVerTS * verticalLength - ((widthVerTS - 2.0 * 5.0 / 1000) * (depthVerTS - 2.0 * 5.0 / 1000) * verticalLength)) * SteelDensityKGPerM;
                double horizontalSteelWeight;

                if (size4Value == 3)
                    horizontalSteelWeight = (widthHorTS * depthHorTS * horizontalLength - ((widthHorTS - 2.0 * 5.0 / 1000) * (depthHorTS - 2.0 * 5.0 / 1000) * horizontalLength)) * SteelDensityKGPerM;
                else
                    horizontalSteelWeight = (widthHorTS * depthHorTS * horizontalLength - ((widthHorTS - 2.0 * 4.0 / 1000) * (depthHorTS - 2.0 * 4.0 / 1000) * horizontalLength)) * SteelDensityKGPerM;

                componentDictionary[HORSTEEL].SetPropertyValue(widthHorTS, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[HORSTEEL].SetPropertyValue(depthHorTS, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[HORSTEEL].SetPropertyValue(horizontalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[HORSTEEL].SetPropertyValue(TSSize1 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0), "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[HORSTEEL].SetPropertyValue(horizontalSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                componentDictionary[VERSTEEL1].SetPropertyValue(widthVerTS, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[VERSTEEL1].SetPropertyValue(depthVerTS, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[VERSTEEL1].SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[VERSTEEL1].SetPropertyValue(TSSize2 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0), "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[VERSTEEL1].SetPropertyValue(verticalSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                componentDictionary[VERSTEEL2].SetPropertyValue(widthVerTS, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[VERSTEEL2].SetPropertyValue(depthVerTS, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[VERSTEEL2].SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[VERSTEEL2].SetPropertyValue(TSSize2 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0), "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[VERSTEEL2].SetPropertyValue(verticalSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateT, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateW, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateD, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", 4 eq spaced " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, holeDiameter, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " dia. holes " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, inset, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " from edges";

                componentDictionary[PLATE1].SetPropertyValue(plateD, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[PLATE1].SetPropertyValue(plateW, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[PLATE1].SetPropertyValue(plateT, "IJOAHgrUtilMetricT", "T");
                componentDictionary[PLATE1].SetPropertyValue(holeDiameter, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                componentDictionary[PLATE1].SetPropertyValue(inset, "IJOAHgrUtilMetricC", "C");
                componentDictionary[PLATE1].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                componentDictionary[PLATE2].SetPropertyValue(plateD, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[PLATE2].SetPropertyValue(plateW, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[PLATE2].SetPropertyValue(plateT, "IJOAHgrUtilMetricT", "T");
                componentDictionary[PLATE2].SetPropertyValue(holeDiameter, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                componentDictionary[PLATE2].SetPropertyValue(inset, "IJOAHgrUtilMetricC", "C");
                componentDictionary[PLATE2].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                double theta2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Z, OrientationAlong.Global_Z);

                //Create graphics
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                {
                    if (Math.Abs(theta2) <= (Math.Atan(1) * 4.0) / 2)
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL, "StartOther", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeZ, 0, -pipeInfo.OutsideDiameter / 2 - shoeH - depthHorTS / 2, X1);
                        JointHelper.CreateRigidJoint(HORSTEEL, "StartOther", VERSTEEL1, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, -depthVerTS / 2, 0, -depthHorTS / 2);
                        JointHelper.CreateRigidJoint(HORSTEEL, "EndOther", VERSTEEL2, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, depthVerTS / 2, 0, -depthHorTS / 2);
                        JointHelper.CreateRigidJoint(VERSTEEL1, "EndOther", PLATE1, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, widthOffsetVerTS + plateT, 0);
                        JointHelper.CreateRigidJoint(VERSTEEL2, "EndOther", PLATE2, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, widthOffsetVerTS + plateT, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL, "StartOther", Plane.XY, Plane.NegativeZX, Axis.X, Axis.Z, 0, pipeInfo.OutsideDiameter / 2 + shoeH + depthHorTS / 2, -X1);
                        JointHelper.CreateRigidJoint(HORSTEEL, "StartOther", VERSTEEL1, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, -depthVerTS / 2, 0, -depthHorTS / 2);
                        JointHelper.CreateRigidJoint(HORSTEEL, "EndOther", VERSTEEL2, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, depthVerTS / 2, 0, -depthHorTS / 2);
                        JointHelper.CreateRigidJoint(VERSTEEL1, "EndOther", PLATE1, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, widthOffsetVerTS + plateT, 0);
                        JointHelper.CreateRigidJoint(VERSTEEL2, "EndOther", PLATE2, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, widthOffsetVerTS + plateT, 0);
                    }
                }
                else
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Support can only be placed on Wall By-Point.", "", "SFS5398_A1.cs", 249);
                    return;
                }
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

                    structConnections.Add(new ConnectionInfo(PLATE1, 1));     //partindex, structindex
                    structConnections.Add(new ConnectionInfo(PLATE2, 1));     //partindex, structindex

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

                BOMString = "Gate support A1 -" + Convert.ToString(sizeBOM) + " SFS 5398";

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
