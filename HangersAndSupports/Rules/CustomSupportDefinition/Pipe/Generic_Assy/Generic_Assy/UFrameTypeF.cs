//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeF.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeF
//   Author       : Hema
//   Creation Date: 17-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Sep-2013    Hema     CR-CP-224494 Converted Generic_Assy to C# .Net 
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

    public class UFrameTypeF : CustomSupportDefinition
    {
        private const string CURVEDPLATE = "CURVEDPLATE";
        private const string TAPERENDPLATE = "TAPERENDPLATE";
        private const string BRACKET1 = "BRACKET1";
        private const string BRACKET2 = "BRACKET2";

        int supportLocation;
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

                    // Get the attributes from assembly
                    supportLocation = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySupLoc", "SupLocation")).PropValue;

                    parts.Add(new PartInfo(CURVEDPLATE, "HgrGen_CurvedPlate_1"));
                    parts.Add(new PartInfo(TAPERENDPLATE, "Util_EndPlTaper_Metric_1"));
                    parts.Add(new PartInfo(BRACKET1, "HgrBracketWithTwoProfile"));
                    parts.Add(new PartInfo(BRACKET2, "HgrBracketWithTwoProfile"));

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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                //Get Pipe OD without insulation
                double pipeDiameter = pipeInfo.OutsideDiameter;

                //Get the Nominal Diameter of the Pipe
                NominalDiameter nominalDiameter = new NominalDiameter();
                nominalDiameter = pipeInfo.NominalDiameter;

                //Get bracket dimensions
                double H, L, L1, W, T, T1, T2;
                H = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeF_Dim", "IJUAHgrGenSrvTypeFDim", "H", "IJUAHgrGenSrvTypeFDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                L = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeF_Dim", "IJUAHgrGenSrvTypeFDim", "L", "IJUAHgrGenSrvTypeFDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                L1 = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeF_Dim", "IJUAHgrGenSrvTypeFDim", "L1", "IJUAHgrGenSrvTypeFDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                W = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeF_Dim", "IJUAHgrGenSrvTypeFDim", "W", "IJUAHgrGenSrvTypeFDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                T = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeF_Dim", "IJUAHgrGenSrvTypeFDim", "T", "IJUAHgrGenSrvTypeFDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                T1 = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeF_Dim", "IJUAHgrGenSrvTypeFDim", "T1", "IJUAHgrGenSrvTypeFDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                T2 = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeF_Dim", "IJUAHgrGenSrvTypeFDim", "T2", "IJUAHgrGenSrvTypeFDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                // '==================================
                //Set properties on part occurrences
                //==================================

                Double angleInDegrees1, angleInDegrees2, angleInRadians1, angleInRadians2, X, Y, Z;

                X = W / 2;
                Y = (pipeDiameter / 2) - T;
                Z = X / Y;

                angleInDegrees1 = 135;
                angleInRadians1 = angleInDegrees1 / 180 * Math.PI;

                angleInDegrees2 = 105;
                angleInRadians2 = angleInDegrees2 / 180 * Math.PI;

                componentDictionary[CURVEDPLATE].SetPropertyValue(angleInRadians1, "IJOAHgrGenCurvedPlate", "Angle");
                componentDictionary[CURVEDPLATE].SetPropertyValue(T, "IJOAHgrGenCurvedPlate", "T");
                componentDictionary[CURVEDPLATE].SetPropertyValue(pipeDiameter / 2, "IJOAHgrGenCurvedPlate", "Radius");
                componentDictionary[CURVEDPLATE].SetPropertyValue(L + 4 * T2, "IJOAHgrGenCurvedPlate", "W");

                componentDictionary[TAPERENDPLATE].SetPropertyValue(pipeDiameter / 2 + T, "IJOAHgrUtilMetricR", "R");
                componentDictionary[TAPERENDPLATE].SetPropertyValue(W, "IJOAHgrUtilMetricW", "W");
                componentDictionary[TAPERENDPLATE].SetPropertyValue(H - T, "IJOAHgrUtilMetricH", "H");
                componentDictionary[TAPERENDPLATE].SetPropertyValue(angleInRadians2, "IJOAHgrUtilMetricAngle", "Angle");
                componentDictionary[TAPERENDPLATE].SetPropertyValue(T2, "IJOAHgrUtilMetricT", "T");

                componentDictionary[BRACKET1].SetPropertyValue(L, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[BRACKET1].SetPropertyValue(H - T, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[BRACKET1].SetPropertyValue(L1, "IJUAHgrBracketOcc", "BottomHeight");
                componentDictionary[BRACKET1].SetPropertyValue(L1, "IJUAHgrBracketOcc", "TopWidth");
                componentDictionary[BRACKET1].SetPropertyValue(T1, "IJUAHgrBracketOcc", "BracketThickness");
                componentDictionary[BRACKET1].SetPropertyValue(0.0001, "IJUAHgrBracketOcc", "CornerRadius");

                componentDictionary[BRACKET2].SetPropertyValue(L, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[BRACKET2].SetPropertyValue(H - T, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[BRACKET2].SetPropertyValue(L1, "IJUAHgrBracketOcc", "BottomHeight");
                componentDictionary[BRACKET2].SetPropertyValue(L1, "IJUAHgrBracketOcc", "TopWidth");
                componentDictionary[BRACKET2].SetPropertyValue(T1, "IJUAHgrBracketOcc", "BracketThickness");
                componentDictionary[BRACKET2].SetPropertyValue(0.0001, "IJUAHgrBracketOcc", "CornerRadius");

                //  '=============
                //Create Joints
                //=============
                if ((SupportHelper.SupportingObjects.Count == 0))
                    //Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    //Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                //Add Joint Between the Curved  Plate and Taper Plate
                JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", TAPERENDPLATE, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -L / 2);

                //Add Joint Between the Curved  Plate and Bracket 1
                JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", BRACKET1, "HgrPort_1", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -pipeDiameter / 2 - T, W / 2, -L / 2 + T2);

                //Add Joint Between the Curved  Plate and Bracket 2
                JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", BRACKET2, "HgrPort_1", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -pipeDiameter / 2 - T, -W / 2 + T1, -L / 2 + T2);
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
                return 1;
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

                    routeConnections.Add(new ConnectionInfo(CURVEDPLATE, 1)); //partindex, routeindex

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

                    //We are not connecting to any structure so we have nothing to return

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
