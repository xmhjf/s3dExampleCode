//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_D1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_D1
//   Author       :  Vijaya
//   Creation Date: 1-Jul-2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1-Jul-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5397_D1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL5397D1 = "HORSTEEL5397_D1";
        private const string VERTSTEEL5397D1 = "VERTSTEEL5397_D1";
        private const int STEELDENSITYKGPERM = 7900;
        private string size5;
        double pipeDiameter = 0.0, shoeHeight = 0.0, overHang = 0.0, horizontalLength;
        PropertyValueCodelist sizeCodeList;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize5", "Size5");
                    size5 = sizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeCodeList.PropValue).DisplayName;

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;

                    horizontalLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal) + overHang;

                    parts.Add(new PartInfo(HORSTEEL5397D1, "Util_Fixed_Box_Metric_1"));
                    if (horizontalLength >= 500.00 / 1000.00)// with clamps                    
                        parts.Add(new PartInfo(VERTSTEEL5397D1, "Utility_CUTBACK_TS1_1"));
                    
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

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                // To get Pipe Nominal Diameter
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                pipeDiameter = pipeInfo.OutsideDiameter + pipeInfo.InsulationThickness * 2;

                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                NominalDiameter minNominalDiameter = new NominalDiameter(), maxNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";

                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_D1.cs", 109);
                    return;
                }

                double maxLengh = 0.0, MA = 0.0, MP = 0.0;
                PropertyValue TSSize1, TSSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size5", size5);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_D1_Dim", "IJUAFINLSrv_SFS5397_D1_Dim", "TS_Size1", parameter, out TSSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_D1_Dim", "IJUAFINLSrv_SFS5397_D1_Dim", "TS_Size2", parameter, out TSSize2);
                string tsSize1 = TSSize1.ToString(), tsSize2 = TSSize2.ToString();


                maxLengh = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D1_Dim", "IJUAFINLSrv_SFS5397_D1_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_D1_Dim", "Size5", size5);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D1_Dim", "IJUAFINLSrv_SFS5397_D1_Dim", "MA", "IJUAFINLSrv_SFS5397_D1_Dim", "Size5", size5);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D1_Dim", "IJUAFINLSrv_SFS5397_D1_Dim", "MP", "IJUAFINLSrv_SFS5397_D1_Dim", "Size5", size5);

                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                double ts1Width, ts1FlangeThickness, ts1WebThickness, ts1Depth, ts2Width, ts2FlangeThickness, ts2WebThickness, ts2Depth; ;


                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", tsSize1, out  ts1Width, out   ts1FlangeThickness, out  ts1WebThickness, out  ts1Depth);
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", tsSize2, out  ts2Width, out   ts2FlangeThickness, out  ts2WebThickness, out  ts2Depth);


                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_D1.cs", 137);
                    return;
                }

                double theta = Math.Abs(Math.PI / 2 - RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "", PortAxisType.Y, OrientationAlong.Global_Z));
                double angleOffset = Math.Tan(theta) * ts1Width;


                double structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                double verticalDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                horizontalLength = structDistance + overHang;


                if (structDistance + overHang > maxLengh)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + (maxLengh * 1000).ToString() + "mm exceeded.", "", "SFS5397_D1.cs", 151);


                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_D1.cs", 155);


                double verticalLength = (550.00 / 1000.00 - 1.0 / 2.0 * (ts2Depth / Math.Tan(Math.PI / 3) + (ts2Depth / Math.Tan(Math.PI / 6))));
                double verticalOffset = 550.00 / 1000.00 * Math.Cos(Math.PI / 6) - horizontalLength;
                double portOrientAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                Matrix4X4 routePortOrientation = new Matrix4X4(), structPortOrientation = new Matrix4X4();
                routePortOrientation = RefPortHelper.PortLCS("Route");
                structPortOrientation = RefPortHelper.PortLCS("Structure");
                string pipePosiion = string.Empty;
                if (Symbols.HgrCompareDoubleService.cmpdbl(structPortOrientation.Origin.Z, routePortOrientation.Origin.Z) == false || Symbols.HgrCompareDoubleService.cmpdbl(structPortOrientation.Origin.X, routePortOrientation.Origin.X) == false || Symbols.HgrCompareDoubleService.cmpdbl(structPortOrientation.Origin.Y, routePortOrientation.Origin.Y) == false)
                {
                    
                    if (structPortOrientation.Origin.Z > routePortOrientation.Origin.Z || structPortOrientation.Origin.X > routePortOrientation.Origin.X || structPortOrientation.Origin.Y > routePortOrientation.Origin.Y)
                        pipePosiion = "BELOW";
                    else
                        pipePosiion = "ABOVE";
                }

                double flipAngle = 0.0;
                int sign = 1;

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

                if ((supportingType == "Slab") && (((supportingType == "Slab") && Symbols.HgrCompareDoubleService.cmpdbl(Math.Abs(portOrientAngle), 0) == true) || ((supportingType == "Steel") && pipePosiion == "ABOVE")))
                {
                    if (Configuration == 1)
                    {
                        flipAngle = Math.PI;
                        sign = -1;
                    }
                    else if (Configuration == 2)
                    {
                        flipAngle = 0.0;
                        sign = 1;
                    }
                }
                else if (Configuration == 2)
                {
                    flipAngle = Math.PI;
                    sign = -1;
                }
                double horSteelWeight = 0.0;

                if (sizeCodeList.PropValue == 3 || sizeCodeList.PropValue == 4)
                    horSteelWeight = (ts1Width * ts1Depth * horizontalLength - ((ts1Width - 2.0 * 5.0 / 1000.00) * (ts1Depth - 2.0 * 5.0 / 1000.00) * horizontalLength)) * STEELDENSITYKGPERM;
                else if (sizeCodeList.PropValue == 5)
                    horSteelWeight = (ts1Width * ts1Depth * horizontalLength - ((ts1Width - 2.0 * 10.0 / 1000.00) * (ts1Depth - 2.0 * 10.0 / 1000.00) * horizontalLength)) * STEELDENSITYKGPERM;
                else
                    horSteelWeight = (ts1Width * ts1Depth * horizontalLength - ((ts1Width - 2.0 * 4.0 / 1000.00) * (ts1Depth - 2.0 * 4.0 / 1000.00) * horizontalLength)) * STEELDENSITYKGPERM;

                string inputBOM = tsSize1 + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL5397D1].SetPropertyValue(ts1Width, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[HORSTEEL5397D1].SetPropertyValue(ts1Depth, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[HORSTEEL5397D1].SetPropertyValue(horizontalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[HORSTEEL5397D1].SetPropertyValue(inputBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[HORSTEEL5397D1].SetPropertyValue(horSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                if (horizontalLength >= 500.00 / 1000.00)
                {

                    inputBOM = tsSize2 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round(verticalLength, 3), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                    componentDictionary[VERTSTEEL5397D1].SetPropertyValue(ts2Width, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    componentDictionary[VERTSTEEL5397D1].SetPropertyValue(ts2Depth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    componentDictionary[VERTSTEEL5397D1].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_L", "L");
                    componentDictionary[VERTSTEEL5397D1].SetPropertyValue(ts2FlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    componentDictionary[VERTSTEEL5397D1].SetPropertyValue(Math.PI / 3, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL5397D1].SetPropertyValue(Math.PI / 6, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL5397D1].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                }

                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (supportingType == "Steel")
                {
                    if (Math.Abs(theta1) < Math.PI / 2)
                    {
                        if (Symbols.HgrCompareDoubleService.cmpdbl(flipAngle, Math.PI) == true)
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, structDistance + overHang, 0, sign * (-pipeDiameter / 2 - ts1Depth / 2 - shoeHeight - angleOffset / 2));
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, structDistance + overHang, 0, sign * (-pipeDiameter / 2 - ts1Depth / 2 - shoeHeight - angleOffset / 2));
                        if (horizontalLength >= 500.00 / 1000.00)
                            JointHelper.CreateRigidJoint(HORSTEEL5397D1, "StartOther", VERTSTEEL5397D1, "EndStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -verticalOffset, 0, ts1Depth / 2);

                    }
                    else    //flip
                    {
                        if (Symbols.HgrCompareDoubleService.cmpdbl(flipAngle, Math.PI) == true)
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, structDistance + overHang, 0, sign * (pipeDiameter / 2 + ts1Depth / 2 + shoeHeight + angleOffset / 2));
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, structDistance + overHang, 0, sign * (pipeDiameter / 2 + ts1Depth / 2 + shoeHeight + angleOffset / 2));

                        if (horizontalLength >= 500.00 / 1000.00)
                            JointHelper.CreateRigidJoint(HORSTEEL5397D1, "StartOther", VERTSTEEL5397D1, "EndStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -verticalOffset, 0, ts1Depth / 2);

                    }
                }
                else
                {
                    if (Symbols.HgrCompareDoubleService.cmpdbl(flipAngle, Math.PI) == true)
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, (structDistance + overHang), -sign * (verticalDistance - pipeDiameter / 2 - ts1Depth / 2 - shoeHeight - angleOffset / 2), 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, (structDistance + overHang), -sign * (verticalDistance - pipeDiameter / 2 - ts1Depth / 2 - shoeHeight - angleOffset / 2), 0);
                    if (horizontalLength >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397D1, "StartOther", VERTSTEEL5397D1, "EndStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -verticalOffset, 0, ts1Depth / 2);

                }

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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

                    routeConnections.Add(new ConnectionInfo(HORSTEEL5397D1, 1));

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

                    structConnections.Add(new ConnectionInfo(HORSTEEL5397D1, 1));

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
            string bomString = "";
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize5", "Size5");
                bomString = "Console support D1 -" + sizeCodeList.PropValue + " SFS 5397";
                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5397_D1" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

