//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5398_H2.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5398_H2
//   Author       :  Manikanth
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Manikanth   CR-CP-224492- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    public class SFS5398_H2 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        string size3;
        double span, shoeH, overHang;
        private const string VERSTEEL = "VerSteel";
        private const string HORSTEEL1 = "HorSteel1";
        private const string HORSTEEL2 = "HorSteel2";
        private const int steelDensityKGperM = 7900;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist sizeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");
                    size3 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;

                    span = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLSpan", "Span")).PropValue;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;

                    parts.Add(new PartInfo(VERSTEEL, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL2, "Util_Fixed_Box_Metric_1"));
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
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                //To get Pipe Nom Dia
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                NominalDiameter currentDiameter = new NominalDiameter();
                currentDiameter = pipeInfo.NominalDiameter;
                double pipeDia = pipeInfo.OutsideDiameter;
                NominalDiameter minNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                NominalDiameter maxNominalDiameter = new NominalDiameter();
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(currentDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5398_H2.cs", 94);
                    return;
                }
                PropertyValue tsSize1, tsSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);

                GenericHelper.GetDataByRule("FINLSrv_SFS5398_H2_Dim", "IJUAFINLSrv_SFS5398_H2_Dim", "TS_Size1", parameter, out tsSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5398_H2_Dim", "IJUAFINLSrv_SFS5398_H2_Dim", "TS_Size2", parameter, out tsSize2);
                double maxSpan = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H2_Dim", "IJUAFINLSrv_SFS5398_H2_Dim", "Max_Span", "IJUAFINLSrv_SFS5398_H2_Dim", "Size3", size3);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H2_Dim", "IJUAFINLSrv_SFS5398_H2_Dim", "Max_Len", "IJUAFINLSrv_SFS5398_H2_Dim", "Size3", size3);
                double ma = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H2_Dim", "IJUAFINLSrv_SFS5398_H2_Dim", "MA", "IJUAFINLSrv_SFS5398_H2_Dim", "Size3", size3);
                double mp = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H2_Dim", "IJUAFINLSrv_SFS5398_H2_Dim", "MP", "IJUAFINLSrv_SFS5398_H2_Dim", "Size3", size3);
                double fMax = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H2_Dim", "IJUAFINLSrv_SFS5398_H2_Dim", "Fmax", "IJUAFINLSrv_SFS5398_H2_Dim", "Size3", size3);

                double widthHorTS, flangeThickness, depthhorTS, webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize1), out widthHorTS, out flangeThickness, out webThickness, out depthhorTS);
                double widthVertTS, depthVertTS, TL;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize2), out widthVertTS, out TL, out webThickness, out depthVertTS);

                double verticalLength = span;
                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5398_H2.cs", 117);
                    return;
                }

                double angleOffset, theta;
                theta = Math.Abs(Math.PI / 2 - RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "", PortAxisType.Y, OrientationAlong.Global_Z));
                angleOffset = Math.Tan(theta) * widthHorTS;
                double verDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                double horDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                double horLength = horDistance + overHang + depthhorTS;

                if (horLength > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + maxLength * 1000 + "mm exceeded.", "", "SFS5398_H2.cs", 129);

                if (span > maxSpan)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable length of " + maxSpan * 1000 + "mm exceeded.", "", "SFS5398_H2.cs", 133);
                    return;
                }

                if (span < pipeDia + shoeH)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Given Shoe Height, span must exceed pipe O.D. plus shoe height: " + pipeDia * 1000 + shoeH * 1000 + " mm.", "", "SFS5398_H2.cs", 138);

                if (overHang < pipeDia / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + pipeDia / 2 * 1000 + " mm.", "", "SFS5398_H2.cs", 1);

                double horSteelWeight, vertSteelWeight;

                horSteelWeight = (widthHorTS * depthhorTS * horLength - ((widthHorTS - 2 * 4.0 / 1000) * (depthhorTS - 2 * 4.0 / 1000) * horLength)) * steelDensityKGperM;

                if (size3 == "1" || size3 == "2")
                    vertSteelWeight = (widthVertTS * depthVertTS * verticalLength - ((widthVertTS - 2 * 4.0 / 1000) * (depthVertTS - 2 * 4.0 / 1000) * verticalLength)) * steelDensityKGperM;
                else
                    vertSteelWeight = (widthVertTS * depthVertTS * verticalLength - ((widthVertTS - 2 * 5.0 / 1000) * (depthVertTS - 2 * 5.0 / 1000) * verticalLength)) * steelDensityKGperM;

                string vertBOM = tsSize1 + "Length" + (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0));

                (componentDictionary[VERSTEEL]).SetPropertyValue(widthHorTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[VERSTEEL]).SetPropertyValue(depthhorTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[VERSTEEL]).SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string hor1BOM = tsSize2 + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[HORSTEEL1]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(horLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(hor1BOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(horSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                (componentDictionary[HORSTEEL2]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(horLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(hor1BOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(horSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                theta = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);
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

                if (supportingType == "Steel")
                {
                    if (Math.Abs(theta) < Math.PI / 2)
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, horDistance + overHang + depthhorTS, 0, -pipeDia / 2 - depthhorTS / 2 - shoeH - angleOffset / 2);
                    else
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, horDistance + overHang + depthhorTS, 0, (pipeDia / 2 + depthhorTS / 2 + shoeH + angleOffset / 2));
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, horDistance + overHang + depthhorTS, (pipeDia / 2 + depthhorTS / 2 + shoeH + angleOffset / 2), 0);
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, horDistance + overHang + depthhorTS, -(pipeDia / 2 + depthhorTS / 2 + shoeH + angleOffset / 2), 0);
                    }
                    else
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Support can only be placed on Slab By-Point. " + pipeDia * 1000 + shoeH * 1000 + " mm.", "", "SFS5398_H2.cs", 202);

                }

                JointHelper.CreateRigidJoint(HORSTEEL1, "StartOther", VERSTEEL, "StartOther", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, depthhorTS / 2, depthhorTS / 2, 0);
                JointHelper.CreateRigidJoint(VERSTEEL, "EndOther", HORSTEEL2, "StartOther", Plane.XY, Plane.ZX, Axis.X, Axis.X, depthhorTS / 2, depthhorTS / 2, 0);


            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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

                    routeConnections.Add(new ConnectionInfo(HORSTEEL1, 1)); // partindex, routeindex


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

                    structConnections.Add(new ConnectionInfo(VERSTEEL, 1)); // partindex, routeindex


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

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            PropertyValueCodelist sizeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLSize3", "Size3");
            size3 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;
            string BOMdesc = "Gate support H2 -" + size3 + " SFS 5398";

            return BOMdesc;
        }
    }
}