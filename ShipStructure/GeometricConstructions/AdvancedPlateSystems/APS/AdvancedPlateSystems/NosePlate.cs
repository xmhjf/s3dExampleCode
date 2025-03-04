//-----------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Component:  NosePlate class is designed to be placed after placement of 
//      standar AC F0012 at the bottom of the Stair Stringer.
//
//      Author:  Alligators
//
//      History:
//      Jan 9, 2015   Alliagtors     CR-262673 New class and Assembly are created.
//-----------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.GeometricConstructions;
using Ingr.SP3D.Common.Middle.GeometricConstructions.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using System;
using Ingr.S3D.Content.Structure.Constants;
using Ingr.S3D.Content.Structure.AdvancedPlateSystems;

namespace Ingr.S3D.Content.Structure
{
    [DefaultLocalizerAttribute(AdvancedPlateSystemsConstants.AdvancedPlateSystemsLocalizer)]
    [IsRemovedAtCommit(false)]

    public class NosePlate : GeometricConstructionDefinition
    {
        #region Define Single input needed by this APS i.e. Selection of StairStringer Part by User

        //Assumption: F0012 Std. AC. is already placed. Otherwise Cancel this APS placement.
        [GraphicInput(AdvancedPlateSystemsConstants.MbrPart, 203, "Select a Stair Stringer (Assumption: F0012 Std. AC. is already placed at End-Frame Connection which is Surface-Default. Otherwise Cancel this APS placement).", "ISPSMemberPartLinear", 1, 1, new string[] { "ISPSPartPrismaticDesignNotify", "ISPSDesignedMemberDesignNotify" })]
        public GraphicInput memberPartGraphicInput;

        #endregion Define Single input needed by this APS i.e. Selection of StairStringer Part by User

        #region Define ComputedInputs

        [ComputedInput(AdvancedPlateSystemsConstants.MbrAxisPort, new string[] { AdvancedPlateSystemsConstants.IJGeometry })]
        public ComputedInput memberAxisPortComputedInput;

        [ComputedInput(AdvancedPlateSystemsConstants.MemberTopFlangeRightBottomPort)]
        public ComputedInput memberTFRBPortComputedInput;

        [ComputedInput(AdvancedPlateSystemsConstants.MbrTopPort)]
        public ComputedInput memberTopPortComputedInput;

        [ComputedInput(AdvancedPlateSystemsConstants.MbrWebRightCurrPort)]
        public ComputedInput memberWebRightCurrPortComputedInput;

        [ComputedInput(AdvancedPlateSystemsConstants.MbrTopRightEndEdgePort)]
        public ComputedInput memberTopRightEndEP27PortComputedInput;

        //Below is an Edge-port between the Top-port and Top-flange-right port of the stringer
        [ComputedInput(AdvancedPlateSystemsConstants.MbrTopIsectTFRightEdgePort)]
        public ComputedInput memberTandTFRightEP33PortComputedInput;

        [ComputedInput(AdvancedPlateSystemsConstants.ADVANCEDPLATESYSTEM, new string[] { AdvancedPlateSystemsConstants.ISPSPartPrismaticDesignNotify, AdvancedPlateSystemsConstants.ISPSDesignedMemberDesignNotify })]
        public ComputedInput GCAdvancedPlateSystem;

        #endregion Define computed inputs

        #region Define Support and Boundary GraphicOutputs

        [GraphicOutput(OutputType.TopologySurface, AdvancedPlateSystemsConstants.Support)]
        public GraphicOutput supportGraphicOutput;

        [GraphicOutput(OutputType.TopologySurface, AdvancedPlateSystemsConstants.Boundary)]
        public GraphicOutput boundaryGraphicOutput;

        #endregion Define Support and Boundary GraphicOutputs

        #region Define Parameters that are Updated during this APS Compute and Shown in Ribbon-bar

        [Parameter(AdvancedPlateSystemsConstants.CornerRadius, 215, AdvancedPlateSystemsConstants.CornerRadius, SP3DPropType.PTDouble, UnitType.Distance, 0.015, true)]
        public Parameter GCcornerRadius;

        [Parameter(AdvancedPlateSystemsConstants.BaseLength, 216, AdvancedPlateSystemsConstants.BaseLength, SP3DPropType.PTDouble, UnitType.Distance, 0.08, true)]
        public Parameter GCbaseLength;

        [Parameter(AdvancedPlateSystemsConstants.Height, 217, AdvancedPlateSystemsConstants.Height, SP3DPropType.PTDouble, UnitType.Distance, 0.05, true)]
        public Parameter GCheight;

        #endregion Define Parameters that are Updated during this APS Compute and Shown in Ribbon-bar

        #region 'Evaluate': Check if Member 'Reflect' Option is True or False and Execute Appropriate Method

        public override void Evaluate()
        {
            MemberPart stairStringerMbrPart = (MemberPart)memberPartGraphicInput.Values[1];
            MemberSystem stairStringerMbrSys = (MemberSystem)stairStringerMbrPart.MemberSystem;

            // Inputs: only one input is needed to place the APS i.e. select a Stair Stringer.

            //Steps involved for creating the APS in evaluate method is broadly explained below 
            // (which is almost common for both 'Reflect' option of Stringer is set either to True or to False):

            //Consider a channel section Stair Stringer: after the placement of the standard AC 
            // (mentioned in the header section of this file) EndCut right-side view looks as below:
            //       _____
            //      |___  |
            //          | |
            //          | |
            //           ¯
            //  APS may be seen as follows:
            //       _____
            //      |___  |
            //       |  | |
            // -->(i)|__| |
            //       (ii)¯
            //Note: since we need stable port to create the APS, we used only the top port of the stringer as 
            // the bounding port (WebbLeft port does not intersect the APS support plane completely
            // for cardinal points from 1-Bottom-Left to the 6-Center-Right).

            //Approach:
            //Extract axis port of the stringer (henceforth termed as memebr) and create a coordinate systsem
            // at the end point of the axis. Using this coordinate system extract the Top-stable-port of
            //the member (which will be used later as bounding port for the APS).

            //Extract the top-flange-right-bottom and the web-right port using current geometry option.
            //Subsequently edge ports are extracted from the two face ports: the length of these two 
            //edge port curves help to determine the start/end point location of edges (i) and (ii) of the APS.

            //Support plane is created with a very small offset into the member from the face created by
            // the web-cut. A plane is constructed at location of edge (ii) acts as one of the boundaries.
            // The other boundary curves i.e. (i) and a fillet arc on right-top portion of the NosePlate
            // are created in plane that is perpendicular to the member axis at a distance from the APS support plane.
            // The boundary curves are linearly extruded to create the remaining two boundary surfaces.

            if (stairStringerMbrSys.Mirror)
            {
                Evaluate_MbrReflectIsTrue();
            }
            else
                Evaluate_MbrReflectIsFalse();

            if (this.Occurrence.GetInputs(GCAdvancedPlateSystem.CollectionName).Count > 0)
                {                
                    foreach (object obj in this.Occurrence.GetInputs(GCAdvancedPlateSystem.CollectionName))
                    {
                        BusinessObject advacedPlateSystemObject = (BusinessObject)obj;                            
                        Type advancedPlateSystemType = advacedPlateSystemObject.GetType();
                        if (advancedPlateSystemType == typeof(string))
                        {
                            string sKey = (string)this.Occurrence.GetInputs(GCAdvancedPlateSystem.CollectionName).GetKey(advacedPlateSystemObject);

                            if (sKey != null && sKey.Equals("JustMigrated"))
                            {
                                this.Occurrence.GetInputs(GCAdvancedPlateSystem.CollectionName).RemoveItemByKey(obj);
                                this.Occurrence.GetInputs(GCAdvancedPlateSystem.CollectionName).Add(advacedPlateSystemObject);
                            }
                        }
                   }
            }
        }

        #endregion 'Evaluate': Check if Member 'Reflect' Option is True or False and Execute Appropriate Method

        #region 'Evaluate_MbrReflectIsTrue': Member Part 'Reflect' Option CheckBox is Checked in the RibbonBar
        //********** Start of 'Evaluate_MbrReflectIsTrue case' method **********
        private void Evaluate_MbrReflectIsTrue()
        {
            try
            {
                #region Obtain APS inputs; Update the APS corner radius parameter value to match with that of the Stringer

                double cornerRadius = (double)GCcornerRadius.Value;
                double baseLength = (double)GCbaseLength.Value;
                double height = (double)GCheight.Value;

                MemberPart stairStringerMbrPart = (MemberPart)memberPartGraphicInput.Values[1];

                //Below code updates corner radius parameter value to that of the standard member radius (to be displayed in ribbon bar)
                double standardMemberRadius = 0;

                standardMemberRadius = AdvancedPlateSystemsCommonFunctions.GetSectionDimensionOfMember(stairStringerMbrPart, AdvancedPlateSystemsConstants.CornerRadius);
                if (standardMemberRadius > AdvancedPlateSystemsConstants.TOLERANCE_VALUE) //we got valid non-zero radius
                    cornerRadius = standardMemberRadius;
                base.Occurrence.SetParameter(AdvancedPlateSystemsConstants.CornerRadius, cornerRadius);

                #endregion Obtain APS inputs; Update the APS corner radius parameter value to match with that of the Stringer

                #region Construct outputs viz. support and boundaries

                //Below code is prepared using anology with VB code generated from the GC menu commands/constructions

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineAxisPortExtractor1 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineAxisPortExtractor");
                lineAxisPortExtractor1.GetInputs("MemberPart").Add(memberPartGraphicInput.Values[1]);
                lineAxisPortExtractor1.Evaluate();
                base.Occurrence.GetInputs(memberAxisPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberAxisPortComputedInput.CollectionName).AddItems(lineAxisPortExtractor1.GetInputs("Port"));


                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd2 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd2.GetInputs("Curve").Add(lineAxisPortExtractor1);
                pointAtCurveEnd2.Evaluate();


                Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem csFromMember3 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem("CSFromMember");
                csFromMember3.GetInputs("MemberPart").Add(memberPartGraphicInput.Values[1]);
                csFromMember3.GetInputs("Point").Add(pointAtCurveEnd2);
                csFromMember3.Evaluate();

                //Current-Internal-Z-Far gives Top-flange-right-bottom port

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor4 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor4.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor4.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor4.SetParameter("GeometrySelector", 2);
                facePortExtractor4.SetParameter("FacesContext", 4);
                facePortExtractor4.SetParameter("LookingAxis", 3);
                facePortExtractor4.SetParameter("IntersectingPlane", 0);
                facePortExtractor4.SetParameter("SurfaceType", 1);
                facePortExtractor4.SetParameter("TrackFlag", 2);
                facePortExtractor4.SetParameter("Offset", 0.0);
                facePortExtractor4.Evaluate();
                base.Occurrence.GetInputs(memberTFRBPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTFRBPortComputedInput.CollectionName).AddItems(facePortExtractor4.GetInputs("Port"));


                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor5 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor5.GetInputs("FacePort").Add(facePortExtractor4);
                edgePortExtractor5.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor5.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor5.SetParameter("GeometrySelector", 4);
                edgePortExtractor5.SetParameter("LookingAxis", 1);
                edgePortExtractor5.SetParameter("TrackFlag", 2);
                edgePortExtractor5.Evaluate();

                //We could optimize by checking if oPointAtCurveStart6 is on oFacePortExtractor8 and avoid below steps
                //for cases where member simple physical aspect does not contain fillet curve. ****
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveStart6 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveStart");
                pointAtCurveStart6.GetInputs("Curve").Add(edgePortExtractor5);
                pointAtCurveStart6.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd7 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd7.GetInputs("Curve").Add(edgePortExtractor5);
                pointAtCurveEnd7.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve8 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve8.GetInputs("Curve").Add(edgePortExtractor5);
                pointAlongCurve8.GetInputs("Point").Add(pointAtCurveEnd7);
                pointAlongCurve8.GetInputs("TrackPoint").Add(pointAtCurveStart6);
                pointAlongCurve8.SetParameter("Distance", 0.5);
                pointAlongCurve8.SetParameter("TrackFlag", 2);
                pointAlongCurve8.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints9 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints9.GetInputs("StartPoint").Add(pointAtCurveStart6);
                lineByPoints9.GetInputs("EndPoint").Add(pointAlongCurve8);
                lineByPoints9.Evaluate();

                //Stable-Lateral-TopPort

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor10 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor10.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor10.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor10.SetParameter("GeometrySelector", 4);
                facePortExtractor10.SetParameter("FacesContext", 2);
                facePortExtractor10.SetParameter("LookingAxis", 3);
                facePortExtractor10.SetParameter("IntersectingPlane", 0);
                facePortExtractor10.SetParameter("SurfaceType", 1);
                facePortExtractor10.SetParameter("TrackFlag", 2);
                facePortExtractor10.SetParameter("Offset", 0);
                facePortExtractor10.Evaluate();

                base.Occurrence.GetInputs(memberTopPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTopPortComputedInput.CollectionName).AddItems(facePortExtractor10.GetInputs("Port"));

                //Current-Lateral-Y-Near gives Web-Right port

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor11 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor11.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor11.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor11.SetParameter("GeometrySelector", 2);
                facePortExtractor11.SetParameter("FacesContext", 2);
                facePortExtractor11.SetParameter("LookingAxis", 2);
                facePortExtractor11.SetParameter("IntersectingPlane", 0);
                facePortExtractor11.SetParameter("SurfaceType", 1);
                facePortExtractor11.SetParameter("TrackFlag", 2);
                facePortExtractor11.SetParameter("Offset", 0);
                facePortExtractor11.Evaluate();

                base.Occurrence.GetInputs(memberWebRightCurrPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberWebRightCurrPortComputedInput.CollectionName).AddItems(facePortExtractor11.GetInputs("Port"));

                //Create pointAlongCurve13 for Member with fillet cases otherwise use pointAtCurveEnd7 reference
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve12 = null;

                double minDist;
                Position posOnline9;
                Position posOnSource;
                ICurve testCurve = (ICurve)lineByPoints9.Output;
                facePortExtractor11.DistanceBetween(testCurve, out minDist, out posOnline9, out posOnSource);
                if (minDist > AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                {
                    pointAlongCurve12 =
                        new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                    pointAlongCurve12.GetInputs("Curve").Add(edgePortExtractor5);
                    pointAlongCurve12.GetInputs("Point").Add(pointAtCurveEnd7);
                    pointAlongCurve12.GetInputs("TrackPoint").Add(pointAtCurveStart6);
                    pointAlongCurve12.SetParameter("Distance", minDist);
                    pointAlongCurve12.SetParameter("TrackFlag", 2);
                    pointAlongCurve12.Evaluate();
                }
                else
                {
                    pointAlongCurve12 = pointAtCurveEnd7;
                }

                //Elevation of below coordinate system is normal to WebRight

                Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem csByPlane13 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem("CSByPlane");
                csByPlane13.GetInputs("Plane").Add(facePortExtractor11);
                csByPlane13.GetInputs("Point").Add(pointAlongCurve12);
                csByPlane13.Evaluate();

                //We know either member coordinate system's Z axis is in same sense of direction of the X axis(local coordinate system)
                // otherwise in opposite direction: this flag helps to flip x, y local coordinate values as needed
                int coorSysFlag = -1;
                if (csFromMember3.ZAxis.Dot(csByPlane13.XAxis) > AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                    coorSysFlag = 1;
                //In VB6, at 15 and 16 indeces we have points inorder to check and set CoorSysFlag, not used here

                //Curr-X-Far gives bottom edge of WebRight port created by WebCut;

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor14 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor14.GetInputs("FacePort").Add(facePortExtractor11);
                edgePortExtractor14.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor14.GetInputs("CoordinateSystem").Add(csByPlane13);
                edgePortExtractor14.SetParameter("GeometrySelector", 2);
                edgePortExtractor14.SetParameter("LookingAxis", 1);
                edgePortExtractor14.SetParameter("TrackFlag", 1);
                edgePortExtractor14.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d oGCPoint;
                Ingr.SP3D.Common.Middle.BusinessObject oGC = pointAlongCurve12.Output;
                oGCPoint = (Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d)oGC;
                Ingr.SP3D.Common.Middle.Point3d  oPoint = new Ingr.SP3D.Common.Middle.Point3d(oGCPoint.Position);
                edgePortExtractor14.DistanceBetween(oPoint, out minDist, out posOnSource);

                //We need to check if the port obtained above is the right one, otherwise get it again
                if (minDist > 0.2 || minDist < 2 * cornerRadius + AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                {
                    edgePortExtractor14 = new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                    edgePortExtractor14.GetInputs("FacePort").Add(facePortExtractor11);
                    edgePortExtractor14.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                    edgePortExtractor14.GetInputs("CoordinateSystem").Add(csByPlane13);
                    edgePortExtractor14.SetParameter("GeometrySelector", 2);
                    edgePortExtractor14.SetParameter("LookingAxis", 1);
                    edgePortExtractor14.SetParameter("TrackFlag", 2);
                    edgePortExtractor14.Evaluate();
                }

                //Align with False case numbering...

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor18 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor18.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor18.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor18.SetParameter("GeometrySelector", 2);
                facePortExtractor18.SetParameter("FacesContext", 2);
                facePortExtractor18.SetParameter("LookingAxis", 3);
                facePortExtractor18.SetParameter("IntersectingPlane", 0);
                facePortExtractor18.SetParameter("SurfaceType", 1);
                facePortExtractor18.SetParameter("TrackFlag", 2);
                facePortExtractor18.SetParameter("Offset", 0);
                facePortExtractor18.Evaluate();


                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor19 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor19.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor19.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor19.SetParameter("GeometrySelector", 2);
                facePortExtractor19.SetParameter("FacesContext", 4);
                facePortExtractor19.SetParameter("LookingAxis", 3);
                facePortExtractor19.SetParameter("IntersectingPlane", 0);
                facePortExtractor19.SetParameter("SurfaceType", 1);
                facePortExtractor19.SetParameter("TrackFlag", 2);
                facePortExtractor19.SetParameter("Offset", 0);
                facePortExtractor19.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveStart20 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveStart");
                pointAtCurveStart20.GetInputs("Curve").Add(edgePortExtractor14);
                pointAtCurveStart20.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints21 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints21.GetInputs("StartPoint").Add(pointAtCurveStart20);
                lineByPoints21.GetInputs("EndPoint").Add(pointAlongCurve12);
                lineByPoints21.Evaluate();

                //To maintain same index value generated in VB code
                double paramDistMeasureLength25 = lineByPoints21.Length;
                int heightInmm = (int)(paramDistMeasureLength25 * AdvancedPlateSystemsConstants.Meter_to_MilliMeter);
                base.Occurrence.SetParameter(AdvancedPlateSystemsConstants.Height, heightInmm * AdvancedPlateSystemsConstants.MilliMeter_to_Meter);

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve22 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve22.GetInputs("Curve").Add(lineByPoints21);
                pointAlongCurve22.GetInputs("Point").Add(pointAlongCurve12);
                pointAlongCurve22.GetInputs("TrackPoint").Add(pointAtCurveStart20);
                pointAlongCurve22.SetParameter("Distance", base.Occurrence.GetParameter("Height"));
                pointAlongCurve22.SetParameter("TrackFlag", 1);
                pointAlongCurve22.Evaluate();


                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor23 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor23.GetInputs("FacePort").Add(facePortExtractor18);
                edgePortExtractor23.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor23.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor23.SetParameter("GeometrySelector", 4);
                edgePortExtractor23.SetParameter("LookingAxis", 1);
                edgePortExtractor23.SetParameter("TrackFlag", 2);
                edgePortExtractor23.Evaluate();
                base.Occurrence.GetInputs(memberTopRightEndEP27PortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTopRightEndEP27PortComputedInput.CollectionName).AddItems(edgePortExtractor23.GetInputs("Port"));

                //Edge between Top and TopFlaneRight

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor24 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor24.GetInputs("FacePort").Add(facePortExtractor18);
                edgePortExtractor24.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor24.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor24.SetParameter("GeometrySelector", 4);
                edgePortExtractor24.SetParameter("LookingAxis", 2);
                edgePortExtractor24.SetParameter("TrackFlag", 1);
                edgePortExtractor24.Evaluate();
                base.Occurrence.GetInputs(memberTandTFRightEP33PortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTandTFRightEP33PortComputedInput.CollectionName).AddItems(edgePortExtractor24.GetInputs("Port"));

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve25 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve25.GetInputs("Curve").Add(edgePortExtractor24);
                pointAlongCurve25.SetParameter("Distance", 0.4);  //400mm
                pointAlongCurve25.SetParameter("TrackFlag", 2);
                pointAlongCurve25.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd26 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd26.GetInputs("Curve").Add(edgePortExtractor23);
                pointAtCurveEnd26.Evaluate();


                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor27 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor27.GetInputs("FacePort").Add(facePortExtractor19);
                edgePortExtractor27.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor27.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor27.SetParameter("GeometrySelector", 4);
                edgePortExtractor27.SetParameter("LookingAxis", 1);
                edgePortExtractor27.SetParameter("TrackFlag", 2);
                edgePortExtractor27.Evaluate();

                double baseLengthFloorVal;
                baseLengthFloorVal = ((int)(edgePortExtractor27.Length * AdvancedPlateSystemsConstants.Meter_to_MilliMeter))
                                            * AdvancedPlateSystemsConstants.MilliMeter_to_Meter; //integer value in mm is converted back to meter units
                base.Occurrence.SetParameter(AdvancedPlateSystemsConstants.BaseLength, baseLengthFloorVal);

                //Considering Channel section for below two variables
                double offsetTFRside = AdvancedPlateSystemsConstants.TOLERANCE_VALUE;
                double offsetWRside = edgePortExtractor27.Length - baseLengthFloorVal;
                if (offsetWRside < AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                    offsetWRside = AdvancedPlateSystemsConstants.TOLERANCE_VALUE;

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve28 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve28.GetInputs("Curve").Add(edgePortExtractor23);
                pointAlongCurve28.GetInputs("Point").Add(pointAtCurveEnd26);
                pointAlongCurve28.SetParameter("Distance", offsetTFRside);
                pointAlongCurve28.SetParameter("TrackFlag", 2);
                pointAlongCurve28.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd29 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd29.GetInputs("Curve").Add(edgePortExtractor24);
                pointAtCurveEnd29.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve30 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve30.GetInputs("Curve").Add(edgePortExtractor24);
                pointAlongCurve30.GetInputs("Point").Add(pointAtCurveEnd26);
                pointAlongCurve30.SetParameter("Distance", 0.3);
                pointAlongCurve30.SetParameter("TrackFlag", 1);
                pointAlongCurve30.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d vectorNormalToSurface31 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("VectorNormalToSurface");
                vectorNormalToSurface31.GetInputs("Surface").Add(facePortExtractor11);
                vectorNormalToSurface31.GetInputs("Point").Add(pointAlongCurve22);
                vectorNormalToSurface31.SetParameter("Range", 1);
                vectorNormalToSurface31.SetParameter("Orientation", 1);
                vectorNormalToSurface31.SetParameter("TrackFlag", 1);
                vectorNormalToSurface31.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints32 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints32.GetInputs("StartPoint").Add(pointAlongCurve25);
                lineByPoints32.GetInputs("EndPoint").Add(pointAtCurveEnd29);
                lineByPoints32.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnCurve33 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnCurve");
                pointByProjectOnCurve33.GetInputs("Point").Add(pointAlongCurve28);
                pointByProjectOnCurve33.GetInputs("Curve").Add(vectorNormalToSurface31);
                pointByProjectOnCurve33.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d planeByPointNormal34 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d("PlaneByPointNormal");
                planeByPointNormal34.GetInputs("Point").Add(pointAlongCurve30);
                planeByPointNormal34.GetInputs("Line").Add(lineAxisPortExtractor1);
                planeByPointNormal34.SetParameter("Range", 2);
                planeByPointNormal34.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnSurf35 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnSurf");
                pointByProjectOnSurf35.GetInputs("Point").Add(pointAlongCurve28);
                pointByProjectOnSurf35.GetInputs("Surface").Add(planeByPointNormal34);
                pointByProjectOnSurf35.GetInputs("Line").Add(lineAxisPortExtractor1);
                pointByProjectOnSurf35.SetParameter("TrackFlag", 1);
                pointByProjectOnSurf35.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve36 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve36.GetInputs("Curve").Add(edgePortExtractor24);
                pointAlongCurve36.GetInputs("Point").Add(pointAtCurveEnd26);
                pointAlongCurve36.GetInputs("TrackPoint").Add(pointAlongCurve30);
                pointAlongCurve36.SetParameter("Distance", AdvancedPlateSystemsConstants.TOLERANCE_VALUE);
                pointAlongCurve36.SetParameter("TrackFlag", 1);
                pointAlongCurve36.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnSurf37 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnSurf");
                pointByProjectOnSurf37.GetInputs("Point").Add(pointByProjectOnCurve33);
                pointByProjectOnSurf37.GetInputs("Surface").Add(planeByPointNormal34);
                pointByProjectOnSurf37.GetInputs("Line").Add(lineAxisPortExtractor1);
                pointByProjectOnSurf37.SetParameter("TrackFlag", 1);
                pointByProjectOnSurf37.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d planeByPointNormal38 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d("PlaneByPointNormal");
                planeByPointNormal38.GetInputs("Point").Add(pointAlongCurve36);
                planeByPointNormal38.GetInputs("Line").Add(edgePortExtractor14);
                planeByPointNormal38.SetParameter("Range", 2);
                planeByPointNormal38.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfFromGType39 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfFromGType");
                surfFromGType39.GetInputs("Surface").Add(planeByPointNormal38);
                surfFromGType39.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints40 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints40.GetInputs("StartPoint").Add(pointByProjectOnSurf37);
                lineByPoints40.GetInputs("EndPoint").Add(pointByProjectOnSurf35);
                lineByPoints40.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve41 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve41.GetInputs("Curve").Add(lineByPoints40);
                pointAlongCurve41.GetInputs("Point").Add(pointByProjectOnSurf35);
                pointAlongCurve41.GetInputs("TrackPoint").Add(pointByProjectOnSurf37);
                pointAlongCurve41.SetParameter("Distance", 0.015); // '15 mm
                pointAlongCurve41.SetParameter("TrackFlag", 2);
                pointAlongCurve41.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints42 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints42.GetInputs("StartPoint").Add(pointByProjectOnSurf37);
                lineByPoints42.GetInputs("EndPoint").Add(pointAlongCurve41);
                lineByPoints42.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve43 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve43.GetInputs("Curve").Add(lineByPoints42);
                pointAlongCurve43.GetInputs("Point").Add(pointByProjectOnSurf37);
                pointAlongCurve43.GetInputs("TrackPoint").Add(pointAlongCurve41);
                pointAlongCurve43.SetParameter("Distance", 1);
                pointAlongCurve43.SetParameter("TrackFlag", 2);
                pointAlongCurve43.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints44 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints44.GetInputs("StartPoint").Add(pointAlongCurve41);
                lineByPoints44.GetInputs("EndPoint").Add(pointAlongCurve43);
                lineByPoints44.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfByLinearExtrusion45 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfByLinearExtrusion");
                surfByLinearExtrusion45.GetInputs("PlanarCrossSection").Add(lineByPoints44);
                surfByLinearExtrusion45.GetInputs("ExtrusionLine").Add(lineByPoints32);
                surfByLinearExtrusion45.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve46 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve46.GetInputs("Curve").Add(lineByPoints21);
                pointAlongCurve46.GetInputs("Point").Add(pointAlongCurve12);
                pointAlongCurve46.GetInputs("TrackPoint").Add(pointAtCurveStart20);
                pointAlongCurve46.SetParameter("Distance", AdvancedPlateSystemsConstants.TOLERANCE_VALUE);
                pointAlongCurve46.SetParameter("TrackFlag", 1);
                pointAlongCurve46.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS47 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS47.GetInputs("CoordinateSystem").Add(csByPlane13);
                pointFromCS47.GetInputs("Point").Add(pointAtCurveStart6);
                pointFromCS47.SetParameter("X", -AdvancedPlateSystemsConstants.TOLERANCE_VALUE * coorSysFlag);
                pointFromCS47.SetParameter("Y", 0);
                pointFromCS47.SetParameter("Z", 0);
                pointFromCS47.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints48 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints48.GetInputs("StartPoint").Add(pointAlongCurve46);
                lineByPoints48.GetInputs("EndPoint").Add(pointFromCS47);
                lineByPoints48.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve49 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve49.GetInputs("Curve").Add(lineByPoints21);
                pointAlongCurve49.GetInputs("Point").Add(pointAlongCurve12);
                pointAlongCurve49.GetInputs("TrackPoint").Add(pointAtCurveStart20);
                pointAlongCurve49.SetParameter("Distance", cornerRadius + AdvancedPlateSystemsConstants.TOLERANCE_VALUE);
                pointAlongCurve49.SetParameter("TrackFlag", 1);
                pointAlongCurve49.Evaluate();

                //Top point of circular arc
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve50 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve50.GetInputs("Curve").Add(lineByPoints48);
                pointAlongCurve50.GetInputs("Point").Add(pointAlongCurve49);
                pointAlongCurve50.SetParameter("Distance", cornerRadius + offsetWRside);
                pointAlongCurve50.SetParameter("TrackFlag", 1);
                pointAlongCurve50.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnSurf51 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnSurf");
                pointByProjectOnSurf51.GetInputs("Point").Add(pointAlongCurve50);
                pointByProjectOnSurf51.GetInputs("Surface").Add(planeByPointNormal34);
                pointByProjectOnSurf51.GetInputs("Line").Add(lineAxisPortExtractor1);
                pointByProjectOnSurf51.SetParameter("TrackFlag", 1);
                pointByProjectOnSurf51.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS52 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS52.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS52.GetInputs("Point").Add(pointByProjectOnSurf51);
                pointFromCS52.SetParameter("X", 0);
                pointFromCS52.SetParameter("Y", 0);
                pointFromCS52.SetParameter("Z", 0.5);
                pointFromCS52.Evaluate();

                //Center
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS53 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS53.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS53.GetInputs("Point").Add(pointByProjectOnSurf51);
                pointFromCS53.SetParameter("X", 0);
                pointFromCS53.SetParameter("Y", 0);
                pointFromCS53.SetParameter("Z", -cornerRadius);
                pointFromCS53.Evaluate();

                //Start
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS54 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS54.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS54.GetInputs("Point").Add(pointFromCS53);
                pointFromCS54.SetParameter("X", 0);
                pointFromCS54.SetParameter("Y", -cornerRadius);
                pointFromCS54.SetParameter("Z", 0);
                pointFromCS54.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS55 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS55.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS55.GetInputs("Point").Add(pointAlongCurve50);
                pointFromCS55.SetParameter("X", 0);
                pointFromCS55.SetParameter("Y", 0);
                pointFromCS55.SetParameter("Z", 0.5);
                pointFromCS55.Evaluate();

                Ingr.SP3D.Common.Middle.Position inPosition = null;
                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints56 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints56.GetInputs("StartPoint").Add(pointAlongCurve50);
                lineByPoints56.GetInputs("EndPoint").Add(pointFromCS55);
                lineByPoints56.Evaluate();
                lineByPoints56.DistanceBetween((ISurface)facePortExtractor10.Output,
                                        out minDist, out posOnSource, out inPosition);

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints57 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints57.GetInputs("StartPoint").Add(pointByProjectOnSurf51);
                lineByPoints57.GetInputs("EndPoint").Add(pointFromCS52);
                lineByPoints57.Evaluate();

                Boolean canIncludeTopPort = true; //initialize
                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints58 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                if (minDist > AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                {
                    canIncludeTopPort = false;
                    Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS59 =
                        new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                    pointFromCS59.GetInputs("CoordinateSystem").Add(csFromMember3);
                    pointFromCS59.GetInputs("Point").Add(pointByProjectOnSurf51);
                    pointFromCS59.SetParameter("X", 0);
                    pointFromCS59.SetParameter("Y", 0.5);
                    pointFromCS59.SetParameter("Z", 0);
                    pointFromCS59.Evaluate();

                    lineByPoints58.GetInputs("StartPoint").Add(pointByProjectOnSurf51);
                    lineByPoints58.GetInputs("EndPoint").Add(pointFromCS59);
                    lineByPoints58.Evaluate();
                }
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS60 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS60.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS60.GetInputs("Point").Add(pointFromCS54);
                pointFromCS60.SetParameter("X", 0);
                pointFromCS60.SetParameter("Y", 0);
                pointFromCS60.SetParameter("Z", -0.5);
                pointFromCS60.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints61 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints61.GetInputs("StartPoint").Add(pointFromCS60);
                lineByPoints61.GetInputs("EndPoint").Add(pointFromCS54);
                lineByPoints61.Evaluate();


                Ingr.SP3D.Common.Middle.GeometricConstructions.Arc3d arcByCenter62 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Arc3d("ArcByCenter");
                arcByCenter62.GetInputs("Center").Add(pointFromCS53);
                arcByCenter62.GetInputs("StartPoint").Add(pointFromCS54);
                arcByCenter62.GetInputs("EndPoint").Add(pointByProjectOnSurf51);
                arcByCenter62.SetParameter("SweepAngle", 1);
                arcByCenter62.SetParameter("TrackFlag", 1);
                arcByCenter62.Evaluate();


                Ingr.SP3D.Common.Middle.GeometricConstructions.ComplexString3d cpxStringByCurves63 =
                new Ingr.SP3D.Common.Middle.GeometricConstructions.ComplexString3d("CpxStringByCurves");
                cpxStringByCurves63.GetInputs("Curves").Add(lineByPoints61, "1");
                cpxStringByCurves63.GetInputs("Curves").Add(arcByCenter62, "2");
                if (canIncludeTopPort)
                    cpxStringByCurves63.GetInputs("Curves").Add(lineByPoints57, "3");
                else
                    cpxStringByCurves63.GetInputs("Curves").Add(lineByPoints58, "3");
                cpxStringByCurves63.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfByLinearExtrusion64 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfByLinearExtrusion");
                surfByLinearExtrusion64.GetInputs("PlanarCrossSection").Add(cpxStringByCurves63);
                surfByLinearExtrusion64.GetInputs("ExtrusionLine").Add(lineByPoints32);
                surfByLinearExtrusion64.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d planeByPointNormal65 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d("PlaneByPointNormal");
                planeByPointNormal65.GetInputs("Point").Add(pointAlongCurve22);
                planeByPointNormal65.GetInputs("Line").Add(lineByPoints21);
                planeByPointNormal65.SetParameter("Range", 0.5);
                planeByPointNormal65.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfFromGType66 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfFromGType");
                surfFromGType66.GetInputs("Surface").Add(planeByPointNormal65);
                surfFromGType66.Evaluate();

                #endregion Construct outputs viz. support and boundaries

                #region Set the outputs viz. support and add only valid boundaries
                GeometricConstructionAssembly thisAssembly = base.Occurrence as GeometricConstructionAssembly;
                if (thisAssembly != null)
                {
                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Support, surfFromGType39.Output, 1);
                    if (canIncludeTopPort)
                        thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, facePortExtractor10.Output,
                                                                     memberTopPortComputedInput.CollectionName);

                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, surfByLinearExtrusion45.Output, 2);
                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, surfFromGType66.Output, 3);
                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, surfByLinearExtrusion64.Output, 4);
                }
                #endregion  Set the outputs viz. 'Support' and add only valid boundaries
            }
            catch(Exception e)
            {                
                throw new Exception("Failed to create NosePlate for Mbr. Reflect = True case");
            }
        }

        #endregion 'Evaluate_MbrReflectIsTrue': Member Part 'Reflect' Option CheckBox is Checked in the RibbonBar
        //********** End of 'Evaluate_MbrReflectIsTrue' method **********

        #region 'Evaluate_MbrReflectIsFalse': Member Part 'Reflect' Option CheckBox is not Checked in the RibbonBar
        //********** Start of 'Evaluate_MbrReflectIsFalse' method **********
        private void Evaluate_MbrReflectIsFalse()
        {
            try
            {
                #region Obtain APS inputs; Update the APS corner radius parameter value to match with that of the Stringer

                double cornerRadius = (double)GCcornerRadius.Value;
                double baseLength = (double)GCbaseLength.Value;
                double height = (double)GCheight.Value;

                MemberPart stairStringerMbrPart = (MemberPart)memberPartGraphicInput.Values[1];

                //Below code updates corner radius parameter value to that of the standard member radius (to be displayed in ribbon bar)
                double standardMemberRadius = 0;

                standardMemberRadius = AdvancedPlateSystemsCommonFunctions.GetSectionDimensionOfMember(stairStringerMbrPart, AdvancedPlateSystemsConstants.CornerRadius);
                if (standardMemberRadius > AdvancedPlateSystemsConstants.TOLERANCE_VALUE) //we got valid non-zero radius
                    cornerRadius = standardMemberRadius;
                base.Occurrence.SetParameter(AdvancedPlateSystemsConstants.CornerRadius, cornerRadius);

                #endregion Obtain APS inputs; Update the APS corner radius parameter value to match with that of the Stringer

                #region Construct outputs viz. support and boundaries

                //Below code is prepared using anology with VB code generated from the GC menu commands/constructions            

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineAxisPortExtractor1 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineAxisPortExtractor");
                lineAxisPortExtractor1.GetInputs("MemberPart").Add(memberPartGraphicInput.Values[1]);
                lineAxisPortExtractor1.Evaluate();
                base.Occurrence.GetInputs(memberAxisPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberAxisPortComputedInput.CollectionName).AddItems(lineAxisPortExtractor1.GetInputs("Port"));

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd2 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd2.GetInputs("Curve").Add(lineAxisPortExtractor1);
                pointAtCurveEnd2.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem csFromMember3 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem("CSFromMember");
                csFromMember3.GetInputs("MemberPart").Add(memberPartGraphicInput.Values[1]);
                csFromMember3.GetInputs("Point").Add(pointAtCurveEnd2);
                csFromMember3.Evaluate();

                //Current-Internal-Z-Far gives Top-flange-right-bottom port

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor4 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor4.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor4.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor4.SetParameter("GeometrySelector", 2);
                facePortExtractor4.SetParameter("FacesContext", 4);
                facePortExtractor4.SetParameter("LookingAxis", 3);
                facePortExtractor4.SetParameter("IntersectingPlane", 0);
                facePortExtractor4.SetParameter("SurfaceType", 1);
                facePortExtractor4.SetParameter("TrackFlag", 2);
                facePortExtractor4.SetParameter("Offset", 0.0);
                facePortExtractor4.Evaluate();
                base.Occurrence.GetInputs(memberTFRBPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTFRBPortComputedInput.CollectionName).AddItems(facePortExtractor4.GetInputs("Port"));

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor5 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor5.GetInputs("FacePort").Add(facePortExtractor4);
                edgePortExtractor5.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor5.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor5.SetParameter("GeometrySelector", 4);
                edgePortExtractor5.SetParameter("LookingAxis", 1);
                edgePortExtractor5.SetParameter("TrackFlag", 2);
                edgePortExtractor5.Evaluate();

                //We could optimize by checking if oPointAtCurveStart6 is on oFacePortExtractor8 and avoid below steps
                //for cases where member simple physical aspect does not contain fillet curve. ****
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveStart6 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveStart");
                pointAtCurveStart6.GetInputs("Curve").Add(edgePortExtractor5);
                pointAtCurveStart6.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd7 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd7.GetInputs("Curve").Add(edgePortExtractor5);
                pointAtCurveEnd7.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve8 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve8.GetInputs("Curve").Add(edgePortExtractor5);
                pointAlongCurve8.GetInputs("Point").Add(pointAtCurveStart6);
                pointAlongCurve8.GetInputs("TrackPoint").Add(pointAtCurveEnd7);
                pointAlongCurve8.SetParameter("Distance", 0.5);
                pointAlongCurve8.SetParameter("TrackFlag", 2);
                pointAlongCurve8.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints9 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints9.GetInputs("StartPoint").Add(pointAtCurveEnd7);
                lineByPoints9.GetInputs("EndPoint").Add(pointAlongCurve8);
                lineByPoints9.Evaluate();

                //Stable-Lateral-TopPort                
                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor10 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor10.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor10.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor10.SetParameter("GeometrySelector", 4);
                facePortExtractor10.SetParameter("FacesContext", 2);
                facePortExtractor10.SetParameter("LookingAxis", 3);
                facePortExtractor10.SetParameter("IntersectingPlane", 0);
                facePortExtractor10.SetParameter("SurfaceType", 1);
                facePortExtractor10.SetParameter("TrackFlag", 2);
                facePortExtractor10.SetParameter("Offset", 0);
                facePortExtractor10.Evaluate();

                base.Occurrence.GetInputs(memberTopPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTopPortComputedInput.CollectionName).AddItems(facePortExtractor10.GetInputs("Port"));

                //Current-Lateral-Y-Near gives Web-Right port

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor11 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor11.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor11.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor11.SetParameter("GeometrySelector", 2);
                facePortExtractor11.SetParameter("FacesContext", 2);
                facePortExtractor11.SetParameter("LookingAxis", 2);
                facePortExtractor11.SetParameter("IntersectingPlane", 0);
                facePortExtractor11.SetParameter("SurfaceType", 1);
                facePortExtractor11.SetParameter("TrackFlag", 1);
                facePortExtractor11.SetParameter("Offset", 0);
                facePortExtractor11.Evaluate();

                base.Occurrence.GetInputs(memberWebRightCurrPortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberWebRightCurrPortComputedInput.CollectionName).AddItems(facePortExtractor11.GetInputs("Port"));
                //Create pointAlongCurve13 for Member with fillet cases otherwise use pointAtCurveStart6 reference

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve12 = null;

                double minDist;
                Position posOnline9;
                Position posOnSource;
                ICurve testCurve = (ICurve)lineByPoints9.Output;
                facePortExtractor11.DistanceBetween(testCurve, out minDist, out posOnline9, out posOnSource);
                if (minDist > AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                {
                    pointAlongCurve12 =
                        new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                    pointAlongCurve12.GetInputs("Curve").Add(edgePortExtractor5);
                    pointAlongCurve12.GetInputs("Point").Add(pointAtCurveStart6);
                    pointAlongCurve12.GetInputs("TrackPoint").Add(pointAtCurveEnd7);
                    pointAlongCurve12.SetParameter("Distance", minDist);
                    pointAlongCurve12.SetParameter("TrackFlag", 2);
                    pointAlongCurve12.Evaluate();
                }
                else
                {
                    pointAlongCurve12 = pointAtCurveStart6;
                }

                //Elevation of below coordinate system is normal to WebRight

                Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem csByPlane13 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.CoordinateSystem("CSByPlane");
                csByPlane13.GetInputs("Plane").Add(facePortExtractor11);
                csByPlane13.GetInputs("Point").Add(pointAlongCurve12);
                csByPlane13.Evaluate();

                //We know either Z axis is in same sense of direction otherwise in opposite direction
                // otherwise in opposite direction: this flag helps to flip x, y local coordinate values as needed
                int coorSysFlag = -1;
                if (csFromMember3.ZAxis.Dot(csByPlane13.XAxis) > AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                    coorSysFlag = 1;
                //In VB6, at 15 and 16 indeces we have points inorder to check and set CoorSysFlag, not used here

                //Curr-X-Far gives bottom edge of WebRight port created by WebCut;                
                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor14 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor14.GetInputs("FacePort").Add(facePortExtractor11);
                edgePortExtractor14.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor14.GetInputs("CoordinateSystem").Add(csByPlane13);
                edgePortExtractor14.SetParameter("GeometrySelector", 2);
                edgePortExtractor14.SetParameter("LookingAxis", 1);
                edgePortExtractor14.SetParameter("TrackFlag", 2);
                edgePortExtractor14.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d oGCPoint;
                Ingr.SP3D.Common.Middle.BusinessObject oGC = pointAlongCurve12.Output;
                oGCPoint = (Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d)oGC;
                Ingr.SP3D.Common.Middle.Point3d oPoint = new Ingr.SP3D.Common.Middle.Point3d(oGCPoint.Position);
                edgePortExtractor14.DistanceBetween(oPoint, out minDist, out posOnSource);

                //We need to check if the port obtained above is the right one, otherwise get it again
                if (minDist > 0.2 || minDist < 2 * cornerRadius + AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                {
                    edgePortExtractor14 = new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                    edgePortExtractor14.GetInputs("FacePort").Add(facePortExtractor11);
                    edgePortExtractor14.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                    edgePortExtractor14.GetInputs("CoordinateSystem").Add(csByPlane13);
                    edgePortExtractor14.SetParameter("GeometrySelector", 2);
                    edgePortExtractor14.SetParameter("LookingAxis", 1);
                    edgePortExtractor14.SetParameter("TrackFlag", 1);
                    edgePortExtractor14.Evaluate();
                }

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveStart14S =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveStart");
                pointAtCurveStart14S.GetInputs("Curve").Add(edgePortExtractor14);
                pointAtCurveStart14S.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd14E =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd14E.GetInputs("Curve").Add(edgePortExtractor14);
                pointAtCurveEnd14E.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPointsOppSense14Line =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPointsOppSense14Line.GetInputs("StartPoint").Add(pointAtCurveEnd14E);
                lineByPointsOppSense14Line.GetInputs("EndPoint").Add(pointAtCurveStart14S);
                lineByPointsOppSense14Line.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor18 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor18.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor18.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor18.SetParameter("GeometrySelector", 2);
                facePortExtractor18.SetParameter("FacesContext", 2);
                facePortExtractor18.SetParameter("LookingAxis", 3);
                facePortExtractor18.SetParameter("IntersectingPlane", 0);
                facePortExtractor18.SetParameter("SurfaceType", 1);
                facePortExtractor18.SetParameter("TrackFlag", 2);
                facePortExtractor18.SetParameter("Offset", 0);
                facePortExtractor18.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface facePortExtractor19 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("FacePortExtractor");
                facePortExtractor19.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                facePortExtractor19.GetInputs("CoordinateSystem").Add(csFromMember3);
                facePortExtractor19.SetParameter("GeometrySelector", 2);
                facePortExtractor19.SetParameter("FacesContext", 4);
                facePortExtractor19.SetParameter("LookingAxis", 3);
                facePortExtractor19.SetParameter("IntersectingPlane", 0);
                facePortExtractor19.SetParameter("SurfaceType", 1);
                facePortExtractor19.SetParameter("TrackFlag", 2);
                facePortExtractor19.SetParameter("Offset", 0);
                facePortExtractor19.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd20 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd20.GetInputs("Curve").Add(edgePortExtractor14);
                pointAtCurveEnd20.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints21 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints21.GetInputs("StartPoint").Add(pointAtCurveEnd20);
                lineByPoints21.GetInputs("EndPoint").Add(pointAlongCurve12);
                lineByPoints21.Evaluate();

                //To maintain same index value generated in VB code
                double paramDistMeasureLength25 = lineByPoints21.Length;
                int heightInmm = (int)(paramDistMeasureLength25 * AdvancedPlateSystemsConstants.Meter_to_MilliMeter);
                base.Occurrence.SetParameter(AdvancedPlateSystemsConstants.Height, heightInmm * AdvancedPlateSystemsConstants.MilliMeter_to_Meter);

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve22 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve22.GetInputs("Curve").Add(lineByPoints21);
                pointAlongCurve22.GetInputs("Point").Add(pointAlongCurve12);
                pointAlongCurve22.GetInputs("TrackPoint").Add(pointAtCurveEnd20);
                pointAlongCurve22.SetParameter("Distance", base.Occurrence.GetParameter("Height"));
                pointAlongCurve22.SetParameter("TrackFlag", 1);
                pointAlongCurve22.Evaluate();


                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor23 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor23.GetInputs("FacePort").Add(facePortExtractor18);
                edgePortExtractor23.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor23.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor23.SetParameter("GeometrySelector", 4);
                edgePortExtractor23.SetParameter("LookingAxis", 1);
                edgePortExtractor23.SetParameter("TrackFlag", 2);
                edgePortExtractor23.Evaluate();
                base.Occurrence.GetInputs(memberTopRightEndEP27PortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTopRightEndEP27PortComputedInput.CollectionName).AddItems(edgePortExtractor23.GetInputs("Port"));

                //Edge between Top and TopFlaneRight                
                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor24 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor24.GetInputs("FacePort").Add(facePortExtractor18);
                edgePortExtractor24.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor24.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor24.SetParameter("GeometrySelector", 4);
                edgePortExtractor24.SetParameter("LookingAxis", 2);
                edgePortExtractor24.SetParameter("TrackFlag", 1);
                edgePortExtractor24.Evaluate();
                base.Occurrence.GetInputs(memberTandTFRightEP33PortComputedInput.CollectionName).Clear();
                base.Occurrence.GetInputs(memberTandTFRightEP33PortComputedInput.CollectionName).AddItems(edgePortExtractor24.GetInputs("Port"));

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve25 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve25.GetInputs("Curve").Add(edgePortExtractor24);
                pointAlongCurve25.SetParameter("Distance", 0.4);  //400mm
                pointAlongCurve25.SetParameter("TrackFlag", 2);
                pointAlongCurve25.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveStart26 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveStart");
                pointAtCurveStart26.GetInputs("Curve").Add(edgePortExtractor23);
                pointAtCurveStart26.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve edgePortExtractor27 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologyCurve("EdgePortExtractor");
                edgePortExtractor27.GetInputs("FacePort").Add(facePortExtractor19);
                edgePortExtractor27.GetInputs("Connectable").Add(memberPartGraphicInput.Values[1]);
                edgePortExtractor27.GetInputs("CoordinateSystem").Add(csFromMember3);
                edgePortExtractor27.SetParameter("GeometrySelector", 4);
                edgePortExtractor27.SetParameter("LookingAxis", 1);
                edgePortExtractor27.SetParameter("TrackFlag", 2);
                edgePortExtractor27.Evaluate();

                double baseLengthFloorVal;
                baseLengthFloorVal = ((int)(edgePortExtractor27.Length * AdvancedPlateSystemsConstants.Meter_to_MilliMeter))
                                            * AdvancedPlateSystemsConstants.MilliMeter_to_Meter; //integer value in mm is converted back to meter units
                base.Occurrence.SetParameter(AdvancedPlateSystemsConstants.BaseLength, baseLengthFloorVal);

                //Considering Channel section for below two variables
                double offsetTFRside = AdvancedPlateSystemsConstants.TOLERANCE_VALUE;
                double offsetWRside = edgePortExtractor27.Length - baseLengthFloorVal;
                if (offsetWRside < AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                    offsetWRside = AdvancedPlateSystemsConstants.TOLERANCE_VALUE;

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve28 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve28.GetInputs("Curve").Add(edgePortExtractor23);
                pointAlongCurve28.GetInputs("Point").Add(pointAtCurveStart26);
                pointAlongCurve28.SetParameter("Distance", offsetTFRside);
                pointAlongCurve28.SetParameter("TrackFlag", 1);
                pointAlongCurve28.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAtCurveEnd29 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAtCurveEnd");
                pointAtCurveEnd29.GetInputs("Curve").Add(edgePortExtractor24);
                pointAtCurveEnd29.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve30 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve30.GetInputs("Curve").Add(edgePortExtractor24);
                pointAlongCurve30.GetInputs("Point").Add(pointAtCurveStart26);
                pointAlongCurve30.SetParameter("Distance", 0.3);
                pointAlongCurve30.SetParameter("TrackFlag", 1);
                pointAlongCurve30.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d vectorNormalToSurface31 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("VectorNormalToSurface");
                vectorNormalToSurface31.GetInputs("Surface").Add(facePortExtractor11);
                vectorNormalToSurface31.GetInputs("Point").Add(pointAlongCurve22);
                vectorNormalToSurface31.SetParameter("Range", 1);
                vectorNormalToSurface31.SetParameter("Orientation", 1);
                vectorNormalToSurface31.SetParameter("TrackFlag", 1);
                vectorNormalToSurface31.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints32 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints32.GetInputs("StartPoint").Add(pointAlongCurve25);
                lineByPoints32.GetInputs("EndPoint").Add(pointAtCurveEnd29);
                lineByPoints32.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnCurve33 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnCurve");
                pointByProjectOnCurve33.GetInputs("Point").Add(pointAlongCurve28);
                pointByProjectOnCurve33.GetInputs("Curve").Add(vectorNormalToSurface31);
                pointByProjectOnCurve33.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d planeByPointNormal34 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d("PlaneByPointNormal");
                planeByPointNormal34.GetInputs("Point").Add(pointAlongCurve30);
                planeByPointNormal34.GetInputs("Line").Add(lineAxisPortExtractor1);
                planeByPointNormal34.SetParameter("Range", 2);
                planeByPointNormal34.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnSurf35 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnSurf");
                pointByProjectOnSurf35.GetInputs("Point").Add(pointAlongCurve28);
                pointByProjectOnSurf35.GetInputs("Surface").Add(planeByPointNormal34);
                pointByProjectOnSurf35.GetInputs("Line").Add(lineAxisPortExtractor1);
                pointByProjectOnSurf35.SetParameter("TrackFlag", 1);
                pointByProjectOnSurf35.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve36 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve36.GetInputs("Curve").Add(edgePortExtractor24);
                pointAlongCurve36.GetInputs("Point").Add(pointAtCurveStart26);
                pointAlongCurve36.GetInputs("TrackPoint").Add(pointAlongCurve30);
                pointAlongCurve36.SetParameter("Distance", AdvancedPlateSystemsConstants.TOLERANCE_VALUE);
                pointAlongCurve36.SetParameter("TrackFlag", 1);
                pointAlongCurve36.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnSurf37 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnSurf");
                pointByProjectOnSurf37.GetInputs("Point").Add(pointByProjectOnCurve33);
                pointByProjectOnSurf37.GetInputs("Surface").Add(planeByPointNormal34);
                pointByProjectOnSurf37.GetInputs("Line").Add(lineAxisPortExtractor1);
                pointByProjectOnSurf37.SetParameter("TrackFlag", 1);
                pointByProjectOnSurf37.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d planeByPointNormal38 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d("PlaneByPointNormal");
                planeByPointNormal38.GetInputs("Point").Add(pointAlongCurve36);
                planeByPointNormal38.GetInputs("Line").Add(lineByPointsOppSense14Line);
                planeByPointNormal38.SetParameter("Range", 2);
                planeByPointNormal38.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfFromGType39 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfFromGType");
                surfFromGType39.GetInputs("Surface").Add(planeByPointNormal38);
                surfFromGType39.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints40 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints40.GetInputs("StartPoint").Add(pointByProjectOnSurf37);
                lineByPoints40.GetInputs("EndPoint").Add(pointByProjectOnSurf35);
                lineByPoints40.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve41 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve41.GetInputs("Curve").Add(lineByPoints40);
                pointAlongCurve41.GetInputs("Point").Add(pointByProjectOnSurf35);
                pointAlongCurve41.GetInputs("TrackPoint").Add(pointByProjectOnSurf37);
                pointAlongCurve41.SetParameter("Distance", 0.015); // '15 mm
                pointAlongCurve41.SetParameter("TrackFlag", 2);
                pointAlongCurve41.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints42 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints42.GetInputs("StartPoint").Add(pointByProjectOnSurf37);
                lineByPoints42.GetInputs("EndPoint").Add(pointAlongCurve41);
                lineByPoints42.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve43 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve43.GetInputs("Curve").Add(lineByPoints42);
                pointAlongCurve43.GetInputs("Point").Add(pointByProjectOnSurf37);
                pointAlongCurve43.GetInputs("TrackPoint").Add(pointAlongCurve41);
                pointAlongCurve43.SetParameter("Distance", 1);
                pointAlongCurve43.SetParameter("TrackFlag", 2);
                pointAlongCurve43.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints44 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints44.GetInputs("StartPoint").Add(pointAlongCurve41);
                lineByPoints44.GetInputs("EndPoint").Add(pointAlongCurve43);
                lineByPoints44.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfByLinearExtrusion45 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfByLinearExtrusion");
                surfByLinearExtrusion45.GetInputs("PlanarCrossSection").Add(lineByPoints44);
                surfByLinearExtrusion45.GetInputs("ExtrusionLine").Add(lineByPoints32);
                surfByLinearExtrusion45.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve46 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve46.GetInputs("Curve").Add(lineByPoints21);
                pointAlongCurve46.GetInputs("Point").Add(pointAlongCurve12);
                pointAlongCurve46.GetInputs("TrackPoint").Add(pointAtCurveEnd20);
                pointAlongCurve46.SetParameter("Distance", AdvancedPlateSystemsConstants.TOLERANCE_VALUE);
                pointAlongCurve46.SetParameter("TrackFlag", 1);
                pointAlongCurve46.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS47 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS47.GetInputs("CoordinateSystem").Add(csByPlane13);
                pointFromCS47.GetInputs("Point").Add(pointAtCurveEnd7);
                pointFromCS47.SetParameter("X", -AdvancedPlateSystemsConstants.TOLERANCE_VALUE * coorSysFlag);
                pointFromCS47.SetParameter("Y", 0);
                pointFromCS47.SetParameter("Z", 0);
                pointFromCS47.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints48 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints48.GetInputs("StartPoint").Add(pointAlongCurve46);
                lineByPoints48.GetInputs("EndPoint").Add(pointFromCS47);
                lineByPoints48.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve49 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve49.GetInputs("Curve").Add(lineByPoints21);
                pointAlongCurve49.GetInputs("Point").Add(pointAlongCurve12);
                pointAlongCurve49.GetInputs("TrackPoint").Add(pointAtCurveEnd20);
                pointAlongCurve49.SetParameter("Distance", cornerRadius + AdvancedPlateSystemsConstants.TOLERANCE_VALUE);
                pointAlongCurve49.SetParameter("TrackFlag", 1);
                pointAlongCurve49.Evaluate();

                //Top point of circular arc
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointAlongCurve50 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointAlongCurve");
                pointAlongCurve50.GetInputs("Curve").Add(lineByPoints48);
                pointAlongCurve50.GetInputs("Point").Add(pointAlongCurve49);
                pointAlongCurve50.SetParameter("Distance", cornerRadius + offsetWRside);
                pointAlongCurve50.SetParameter("TrackFlag", 1);
                pointAlongCurve50.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointByProjectOnSurf51 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointByProjectOnSurf");
                pointByProjectOnSurf51.GetInputs("Point").Add(pointAlongCurve50);
                pointByProjectOnSurf51.GetInputs("Surface").Add(planeByPointNormal34);
                pointByProjectOnSurf51.GetInputs("Line").Add(lineAxisPortExtractor1);
                pointByProjectOnSurf51.SetParameter("TrackFlag", 1);
                pointByProjectOnSurf51.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS52 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS52.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS52.GetInputs("Point").Add(pointByProjectOnSurf51);
                pointFromCS52.SetParameter("X", 0);
                pointFromCS52.SetParameter("Y", 0);
                pointFromCS52.SetParameter("Z", 0.5);
                pointFromCS52.Evaluate();

                //Center
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS53 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS53.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS53.GetInputs("Point").Add(pointByProjectOnSurf51);
                pointFromCS53.SetParameter("X", 0);
                pointFromCS53.SetParameter("Y", 0);
                pointFromCS53.SetParameter("Z", -cornerRadius);
                pointFromCS53.Evaluate();

                //Start
                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS54 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS54.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS54.GetInputs("Point").Add(pointFromCS53);
                pointFromCS54.SetParameter("X", 0);
                pointFromCS54.SetParameter("Y", cornerRadius);
                pointFromCS54.SetParameter("Z", 0);
                pointFromCS54.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS55 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS55.GetInputs("CoordinateSystem").Add(csByPlane13);
                pointFromCS55.GetInputs("Point").Add(pointAlongCurve50);
                pointFromCS55.SetParameter("X", (AdvancedPlateSystemsConstants.TOLERANCE_VALUE + 0.5) * coorSysFlag);
                pointFromCS55.SetParameter("Y", 0);
                pointFromCS55.SetParameter("Z", 0);
                pointFromCS55.Evaluate();

                Ingr.SP3D.Common.Middle.Position inPosition = null;
                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints56 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints56.GetInputs("StartPoint").Add(pointAlongCurve50);
                lineByPoints56.GetInputs("EndPoint").Add(pointFromCS55);
                lineByPoints56.Evaluate();
                lineByPoints56.DistanceBetween((ISurface)facePortExtractor10.Output,
                                        out minDist, out posOnSource, out inPosition);

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints57 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints57.GetInputs("StartPoint").Add(pointByProjectOnSurf51);
                lineByPoints57.GetInputs("EndPoint").Add(pointFromCS52);
                lineByPoints57.Evaluate();

                Boolean canIncludeTopPort = true; //initialize
                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints58 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                if (minDist > AdvancedPlateSystemsConstants.TOLERANCE_VALUE)
                {
                    canIncludeTopPort = false;
                    Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS59 =
                        new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                    pointFromCS59.GetInputs("CoordinateSystem").Add(csFromMember3);
                    pointFromCS59.GetInputs("Point").Add(pointByProjectOnSurf51);
                    pointFromCS59.SetParameter("X", 0);
                    pointFromCS59.SetParameter("Y", -0.5);
                    pointFromCS59.SetParameter("Z", 0);
                    pointFromCS59.Evaluate();

                    lineByPoints58.GetInputs("StartPoint").Add(pointByProjectOnSurf51);
                    lineByPoints58.GetInputs("EndPoint").Add(pointFromCS59);
                    lineByPoints58.Evaluate();
                }

                Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d pointFromCS60 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Point3d("PointFromCS");
                pointFromCS60.GetInputs("CoordinateSystem").Add(csFromMember3);
                pointFromCS60.GetInputs("Point").Add(pointFromCS54);
                pointFromCS60.SetParameter("X", 0);
                pointFromCS60.SetParameter("Y", 0);
                pointFromCS60.SetParameter("Z", -0.5);
                pointFromCS60.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d lineByPoints61 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Line3d("LineByPoints");
                lineByPoints61.GetInputs("StartPoint").Add(pointFromCS60);
                lineByPoints61.GetInputs("EndPoint").Add(pointFromCS54);
                lineByPoints61.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Arc3d arcByCenter62 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Arc3d("ArcByCenter");
                arcByCenter62.GetInputs("Center").Add(pointFromCS53);
                arcByCenter62.GetInputs("StartPoint").Add(pointFromCS54);
                arcByCenter62.GetInputs("EndPoint").Add(pointByProjectOnSurf51);
                arcByCenter62.SetParameter("SweepAngle", 1);
                arcByCenter62.SetParameter("TrackFlag", 1);
                arcByCenter62.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.ComplexString3d cpxStringByCurves63 =
                new Ingr.SP3D.Common.Middle.GeometricConstructions.ComplexString3d("CpxStringByCurves");
                cpxStringByCurves63.GetInputs("Curves").Add(lineByPoints61, "1");
                cpxStringByCurves63.GetInputs("Curves").Add(arcByCenter62, "2");
                if (canIncludeTopPort)
                    cpxStringByCurves63.GetInputs("Curves").Add(lineByPoints57, "3");
                else
                    cpxStringByCurves63.GetInputs("Curves").Add(lineByPoints58, "3");
                cpxStringByCurves63.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfByLinearExtrusion64 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfByLinearExtrusion");
                surfByLinearExtrusion64.GetInputs("PlanarCrossSection").Add(cpxStringByCurves63);
                surfByLinearExtrusion64.GetInputs("ExtrusionLine").Add(lineByPoints32);
                surfByLinearExtrusion64.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d planeByPointNormal65 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.Plane3d("PlaneByPointNormal");
                planeByPointNormal65.GetInputs("Point").Add(pointAlongCurve22);
                planeByPointNormal65.GetInputs("Line").Add(lineByPoints21);
                planeByPointNormal65.SetParameter("Range", 0.5);
                planeByPointNormal65.Evaluate();

                Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface surfFromGType66 =
                    new Ingr.SP3D.Common.Middle.GeometricConstructions.TopologySurface("SurfFromGType");
                surfFromGType66.GetInputs("Surface").Add(planeByPointNormal65);
                surfFromGType66.Evaluate();

                #endregion Construct outputs viz. support and boundaries

                #region Set the outputs viz. support and add only valid boundaries


                GeometricConstructionAssembly thisAssembly = base.Occurrence as GeometricConstructionAssembly;
                if (thisAssembly != null)
                {
                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Support, surfFromGType39.Output, 1);
                    if (canIncludeTopPort)
                        thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, facePortExtractor10.Output, memberTopPortComputedInput.CollectionName);
                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, surfByLinearExtrusion45.Output, 2);
                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, surfFromGType66.Output, 3);
                    thisAssembly.SetOutput(AdvancedPlateSystemsConstants.Boundary, surfByLinearExtrusion64.Output, 4);
                }
                //********** End of 'Evaluate_MbrReflectIsFalse' method **********
                #endregion  Set the outputs viz. 'Support' and add only valid boundaries
            }
            catch(Exception)
            {
                throw new Exception("Failed to create NosePlate for Mbr. Reflect = False case");
            }
        }
        #endregion 'Evaluate_MbrReflectIsFalse': Member Part 'Reflect' Option CheckBox is not Checked in the RibbonBar
        //********** End of 'Evaluate_MbrReflectIsFalse' method **********
    }
}
