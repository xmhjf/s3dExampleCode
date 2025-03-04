//-----------------------------------------------------------------------------
//      Copyright (C) 2010-14 Intergraph Corporation.  All rights reserved.
//
//      Component:  S3DEdgeMap class is designed to call the appropriate
//                  cross section class based on the section type.  
//
//      Author:  3XCalibur
//
//      History:
//      November 03, 2010   WR              Created
//      July 21, 2011       Alligators      CR-CP-199876 Expose a rule for calculating the sketching plane for a tube end cut  
//      September 12,2011   Alliagtors      TR-CP-202226 Enabled Genric AC's for Tube(B'ded) Vs Std Mbr(B'ding).
//                                          Modified GetSketchPlaneForTube() to take care if the Generic AC calls this method
//      October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) (i) 'GetSectionEdges'is added with new parameter 'pointMap' 
//                                          (ii)'GetCrossSectionMap' is called with additional parameter 'pointMap'.
//      April 23,2013   Alliagtors          TR-230220 Updated Generic AC for Tube(B'ded) Vs Stiffener(B'ding).
//                                          Modified GetSketchPlaneForTube() to take care it. Also, IsPenetrationWithSketchingPlaneNormal() is 
//                                          modified to return false for the above case.
//      November 03,2014    MDT/GH          CR-250198 Lapped AC for traffic items	
//		November 02,2015	knukala				DI-CP-275245,275246,275247,275248 Removed Introp calls
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Content.Structure.CrossSectionMappings;
using Ingr.SP3D.Common.Middle;
using System;

using Ingr.SP3D.Common.Middle.Services.Hidden;
using System.Collections.ObjectModel; 


namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// The S3DEdgeMap class implements all the logic to do the following:
    /// 1. Return a cross section edge map based on the type of cross section.
    /// 2. Get the correct orientation quadrant.
    /// 3. Get the Web penetrating flag.
    /// </summary>
    internal class S3DEdgeMap
    {       
        private IPort    boundingPort;
        private IPort    boundedPort;
        private Position boundedLocation;

        /// <summary>
        /// Initializes a new instance of the <see cref="S3DEdgeMap"/> class.
        /// </summary>
        /// <param name="boundingPort">The bounding port.</param>
        /// <param name="boundedPort">The bounded port.</param>
        /// <exception cref="ArgumentNullException">The passed port is null.</exception>
        internal S3DEdgeMap(IPort boundingPort, IPort boundedPort)
        {
            if (boundingPort == null)
            {
                throw new ArgumentNullException("boundingPort");
            }

            if (boundedPort == null)
            {
                throw new ArgumentNullException("boundedPort");
            }

            this.boundingPort = boundingPort;
            this.boundedPort = boundedPort;

            // Get the bounded port position
            Position boundedStart;
            Position boundedEnd;

            MemberPartAxisPort memberAxisPort = this.boundedPort as MemberPartAxisPort;
            MemberPart memberObj = this.boundedPort.Connectable as MemberPart;
            StiffenerPartBase StiffObj = this.boundedPort.Connectable as StiffenerPartBase;
            TopologyPort oTopoPort = this.boundedPort as TopologyPort;

            if (memberObj != null)
            {
                if (memberAxisPort == null)
                {
                    // To do - declare a new exception ?
                }
               // GetEndPoints method call raises an exception for curved member part,
               // so use an alternative approach
               // .... memberObj.GetEndPoints(out boundedStart, out boundedEnd);

                IPoint boundedStPos = memberObj.GetPointAtEnd(MemberAxisEnd.Start);
                // Initialize the variables.
                double x = 0.0, y = 0.0, z = 0.0;
                x= boundedStPos.X;
                y = boundedStPos.Y;
                z = boundedStPos.Z;
                boundedStart = new Position();
                boundedStart.Set (x, y, z);

                IPoint boundedEnPos = memberObj.GetPointAtEnd(MemberAxisEnd.End);
                // Initialize the variables.
                x = boundedEnPos.X;
                y = boundedEnPos.Y;
                z = boundedEnPos.Z;
                boundedEnd = new Position();
                boundedEnd.Set(x, y, z);

                if (memberAxisPort != null && memberAxisPort.AxisPortType == MemberAxisPortType.Start)
                {
                    this.boundedLocation = boundedStart;
                }
                else if (memberAxisPort != null && memberAxisPort.AxisPortType == MemberAxisPortType.End)
                {
                    this.boundedLocation = boundedEnd;
                }
                else
                {

                    if (memberAxisPort != null && memberAxisPort.AxisPortType == MemberAxisPortType.Along)
                    {
                        MemberPartAxisPort BdgmemberAxisPort = this.boundingPort as MemberPartAxisPort;
                        if (BdgmemberAxisPort != null && BdgmemberAxisPort.AxisPortType == MemberAxisPortType.Along)
                        {
                            MemberPart boundingMemberPart = this.boundingPort.Connectable as MemberPart;
                            ICurve boundingAxisCurve = boundingMemberPart.Axis;
                            ICurve boundedAxisCurve = memberObj.Axis;
                            Position boundedAlongPos = new Position();
                            Position boundingAlongPos = new Position();
                            double minDist;
                            boundedAxisCurve.DistanceBetween(boundingAxisCurve, out minDist, out boundedAlongPos, out boundingAlongPos);
                            this.boundedLocation = boundedAlongPos;
                        }
                    }
                }
            }
            else if (StiffObj != null)
            {
                StiffenerSystemBase StiffenerSystem = StiffObj.SystemParent as StiffenerSystemBase;

                if (StiffenerSystem == null)
                {
                    throw new ArgumentNullException("StiffenerSystem Null");
                }

                Curve3d StiffenerAxis = StiffenerSystem.Axis as Curve3d;

                if (oTopoPort == null)
                {
                    // To do - declare a new exception ?
                }

                boundedStart = new Position();
                boundedEnd = new Position();
                StiffenerAxis.EndPoints(out boundedStart,out boundedEnd);
                Position BoundedPos = new Position();
                this.boundedLocation = new Position();
                if (oTopoPort != null && oTopoPort.ContextId == ContextTypes.Base)
                {
                    this.boundedLocation.Set(boundedStart.X, boundedStart.Y, boundedStart.Z);
                }
                else
                {
                    this.boundedLocation.Set(boundedEnd.X, boundedEnd.Y, boundedEnd.Z);
                }
            }
            else
            {
                // To do - declare a new exception
            }
        }

        /// <summary>
        /// Gets the section edges.
        /// </summary>
        /// <param name="sectionType">Type of the section.</param>
        /// <param name="flipLeftAndRight">if set to <c>true</c> then left and right will be swapped.</param>
        /// <param name="quadrant">The quadrant.</param>
        /// <returns>The cross section mapping of edges based on the section type, quadrant, and flip flag.</returns>
        /// <exception cref="ArgumentNullException">Argument sectionType is null.</exception>
        internal Dictionary<SectionFaceType, SectionFaceType> GetSectionEdges(string sectionType, bool flipLeftAndRight, int quadrant, Dictionary<int, int> pointMap)
        {
            if (sectionType == null)
            {
                throw new ArgumentNullException("sectionType");
            }

            Dictionary<SectionFaceType, SectionFaceType> mappedEdges = new Dictionary<SectionFaceType, SectionFaceType>();
            ICrossSectionMap crossSectionMap = null;

            switch (sectionType)
            {
                case "2L":
                    crossSectionMap = new Section2LMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "C":
                    crossSectionMap = new SectionCMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "HP":
                    //similar to W use the same map
                    crossSectionMap = new SectionWMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "HSSC":
                    crossSectionMap = new SectionCircularMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "CS":
                    crossSectionMap = new SectionSolidCircularMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "HSSR":
                    crossSectionMap = new SectionRectangularMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

				case "RS":
                    crossSectionMap = new SectionRectangularMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;
                case "L":
                    crossSectionMap = new SectionLMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "M":
                    //similar to W use the same map
                    crossSectionMap = new SectionWMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "MC":
                    //similar to C use the same map
                    crossSectionMap = new SectionCMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "MT":
                    //similar to T use the same map
                    crossSectionMap = new SectionTMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "PIPE":
                    crossSectionMap = new SectionCircularMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;
                case "S":
                    //similar to W use the same map
                    crossSectionMap = new SectionWMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "ST":
                    //similar to T use the same map
                    crossSectionMap = new SectionTMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "W":
                    crossSectionMap = new SectionWMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "WT":
                    //similar to T use the same map
                    crossSectionMap = new SectionTMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "EA":
                    //EA Profile Section Type
                    crossSectionMap = new SectionEAUAMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "UA":
                    //Similar to EA, use the same map
                    crossSectionMap = new SectionEAUAMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "FB":
                    //Flat Bar
                    crossSectionMap = new SectionFBMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "I":
                    //similar to W use the same map
                    crossSectionMap = new SectionWMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "H":
                    //similar to W use the same map
                    crossSectionMap = new SectionWMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "ISType":
                    //similar to W use the same map
                    crossSectionMap = new SectionWMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "T_XType":
                    //similar to T use the same map
                    crossSectionMap = new SectionTMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "TSType":
                    //similar to T use the same map
                    crossSectionMap = new SectionTMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "BUT":
                    //similar to T use the same map
                    crossSectionMap = new SectionTMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "BUTL2":
                    //similar to T use the same map
                    crossSectionMap = new SectionTMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "BUTL3":
                    //BUTL3 Type
                    crossSectionMap = new SectionBUTL3Map();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "C_SS":
                    //similar to C use the same map
                    crossSectionMap = new SectionCMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                case "CSType":
                    //similar to C use the same map
                    crossSectionMap = new SectionCMap();
                    mappedEdges = crossSectionMap.GetCrossSectionMap(flipLeftAndRight, quadrant, pointMap);
                    break;

                default:
                    throw new EdgeMappingException("Not a valid cross section type.");
            }

            return mappedEdges;
        }

        /// <summary>
        /// Returns a local coordinate system useful for end cut logic
        /// </summary>
        /// <param name="memberPort">The bounded or bounding port.</param>
        /// <param name="uAxis">Local x-direction, guaranteed to equal the web-right normal (web penetrated) or top normal (flange penetrated).</param>
        /// <param name="vAxis">Local y-direction; the direction from bottom to top (web penetrated) or right to left (flange penetrated).</param>
        /// <param name="wAxis">Local z-direction; the axis direction, adjusted to point away from the end port if this is the bounded object</param>
        /// <returns>void</returns>
        /// <exception cref="ArgumentNullException">The passed port is null.</exception>
        internal void GetEndCutLocalCoordinateSystem(IPort inputPort, out Vector uAxis, out Vector vAxis, out Vector wAxis)
        {
            if (inputPort == null)
            {
                throw new ArgumentNullException("inputPort");
            }

            // ----------------------------------------------------------
            // If bounded object, get the sketching plane if it is a tube
            // ----------------------------------------------------------
            Plane3d sketchPlane = null;
            if (inputPort == boundedPort)
            {
                MemberPart boundedMember = inputPort.Connectable as MemberPart;
                if (IsMemberATubeType(boundedMember) == true)
                {
                    sketchPlane = GetSketchPlaneForTube();
                }
            }

            MemberPartAxisPort memberAxisPort = inputPort as MemberPartAxisPort;
            MemberPart memberObj = inputPort.Connectable as MemberPart;
            StiffenerPartBase StiffenerObj = inputPort.Connectable as StiffenerPartBase;

            // Only members/Stiffeners are supported for now
            if (memberObj == null && StiffenerObj == null)
            {
                throw new ArgumentNullException("inputPort");
            }

            TopologyPort oTopoPort = inputPort as TopologyPort;

            //if Port connectable is member Obj and Port is not MemberAxisPort
            //then try to get Member Axis Port from Member Obj
            if (memberObj != null && memberAxisPort == null)
            {
                MemberAxisPortType eAxisPort = MemberAxisPortType.Along;

                if (oTopoPort != null)
                {
                    if (oTopoPort.ContextId == ContextTypes.Base)
                    {
                        eAxisPort = MemberAxisPortType.Start;
                    }
                    else if (oTopoPort.ContextId == ContextTypes.Offset)
                    {
                        eAxisPort = MemberAxisPortType.End;
                    }
                }

                memberAxisPort = (MemberPartAxisPort)memberObj.GetAxisPort(eAxisPort);    
            }

            // --------------------------------------
            // u-vector
            // --------------------------------------
            // If the sketching plane was calculated for a tube, use it
            // get the member y-axis and convert it to unit-vector
            Matrix4X4 oMemberMatrix = new Matrix4X4();
            if (memberObj != null)
            {
                oMemberMatrix = memberObj.GetMatrixAtPosition(this.boundedLocation);
            }

            Matrix4X4 oStiffenerMatrix = new Matrix4X4();
            if (StiffenerObj != null)
            {
                oStiffenerMatrix = StiffenerObj.GetMatrixAtPosition(this.boundedLocation);
            }

            if (sketchPlane != null)
            {
                uAxis = new Vector(sketchPlane.Normal);
            }
            else if (memberObj != null)
            {
                if (memberObj.Curved)
            {
                 // Set u, v and w for curved 
                    // u = Y Axis of Member
                uAxis = new Vector();
                    uAxis.Set(oMemberMatrix.GetIndexValue(4), oMemberMatrix.GetIndexValue(5), oMemberMatrix.GetIndexValue(6));
            }
            else
            {
                uAxis = new Vector(memberObj.YAxis);
            }
            }
            else if (StiffenerObj != null)
            {
                //Get u Vector From Matrix
                uAxis = new Vector();
                uAxis.Set(oStiffenerMatrix.GetIndexValue(0), oStiffenerMatrix.GetIndexValue(1), oStiffenerMatrix.GetIndexValue(2));
            }
            else
            {
                throw new ArgumentNullException();   
            }
            uAxis.Length = 1.0;


            // We'll want to convert to the following for all three vectors, once we need to support curved members
            //// Get the web-right port geometry
            //TopologyPort webRightPort = memberObj.GetPort(TopologyGeometryType.Face, 2000, 258, ContextTypes.Lateral, 258);
            //Surface3d    webRightGeom = (Surface3d)webRightPort.Geometry;

            //// Project the bounded location onto the port and evaluate the normal
            //Position closestPoint = webRightGeom.ProjectPoint(this.boundedLocation);
            //uAxis = webRightGeom.OutwardNormalAtPoint(closestPoint);
            //uAxis.Length = 1.0;

            // --------------------------------------
            // v-vector
            // --------------------------------------
            // If the sketching plane was calculated for a tube, use it
            // Otherwise, get the member z-axis and convert it to unit-vector
            if (sketchPlane != null)
            {
                vAxis =  new Vector(sketchPlane.VDirection);
            }
            else if (memberObj != null)
            {
                if (memberObj.Curved)
            {
                    //v = Z Axis of Member
                vAxis = new Vector();
                vAxis.Set(oMemberMatrix.GetIndexValue(8), oMemberMatrix.GetIndexValue(9), oMemberMatrix.GetIndexValue(10));
            }
            else
            {
                vAxis = new Vector(memberObj.ZAxis);
            }
            }
            else if (StiffenerObj != null)
            {
                //Get v Vector from Matrix
                vAxis = new Vector();
                vAxis.Set(oStiffenerMatrix.GetIndexValue(4), oStiffenerMatrix.GetIndexValue(5), oStiffenerMatrix.GetIndexValue(6));
            }
            else
            {
                throw new ArgumentNullException();
            }
            vAxis.Length = 1.0;

            // --------------------------------------
            // w-vector
            // --------------------------------------
            // If the sketching plane was calculated for a tube, use it
            // Otherwise, get the member x-axis and convert it to unit-vector
            if (sketchPlane != null)
            {
                // The sketching plane points toward the part, so we need to reverse it.
                wAxis = new Vector(-sketchPlane.UDirection.X, -sketchPlane.UDirection.Y, -sketchPlane.UDirection.Z);
            }
            else if (memberObj != null)
            {
                if (memberObj.Curved)
            {
                    //w = X Axis of Member
                wAxis = new Vector();
                wAxis.Set(oMemberMatrix.GetIndexValue(0), oMemberMatrix.GetIndexValue(1), oMemberMatrix.GetIndexValue(2));
                wAxis.Length = 1.0;
                // Negate the u-vector if the port is at the start, so that it points away from the member
                if (memberAxisPort != null && memberAxisPort.AxisPortType == MemberAxisPortType.Start)
                {
                    wAxis.Length = -1.0;
                }
            }
            else
            {
                wAxis = new Vector(memberObj.XAxis);
                wAxis.Length = 1.0;
          
                // Negate the u-vector if the port is at the start, so that it points away from the member
                if((memberAxisPort != null && memberAxisPort.AxisPortType == MemberAxisPortType.Start) || 
                        (memberAxisPort != null && memberAxisPort.AxisPortType == MemberAxisPortType.Along && inputPort.Equals(boundedPort)))
                {
                    wAxis.Length = -1.0;
                }
            }
            }
            else if (StiffenerObj != null)
            {
                // Apparently the w-vector of a member axis is constant even as the member is flipped, but the stiffener matrix
                // w-vector flips with the secondary orientation (always W = U x V)
                // We will get the landing curve tangent and reverse it if we are at the base port
                wAxis = new Vector();
                
                StiffenerSystemBase stiffSys = StiffenerObj.SystemParent as StiffenerSystemBase;
                if (stiffSys != null)
                {
                    Curve3d landingCurve = stiffSys.Axis;
                    wAxis = landingCurve.TangentAtPoint(boundedLocation);
                    wAxis.Length = 1.0;
                    
                    if (oTopoPort != null && oTopoPort.ContextId == ContextTypes.Base)
                    {
                        wAxis.Length = -1.0;
                    }
                }
                else // standalone stiffeners not supported by this method
                {
                    throw new ArgumentNullException();
                }
            }
            else
            {
                throw new ArgumentNullException();
            }

            // -----------------------------------------------------------------------------------------------
            // Reverse u unless mirror flag is set; Reverse w if mirror is set and this is the bounding object
            // Do not reverse either if this is a bounded tube
            // -----------------------------------------------------------------------------------------------
            // if the sketch plane is null, this is not a tube
            if ((sketchPlane == null) || inputPort.Equals(boundingPort))
            {
                if (memberObj != null)
                {
                    if (!memberObj.Mirror)
                        uAxis.Length = -1.0;
                    else if (inputPort.Equals(boundingPort))
                        wAxis.Length = -1.0;
                }
                else if (StiffenerObj != null)
                {
                    //If the profile is the bounding object and the profile coordinate system is right handed
                    //then we need to negate w so that we preserve a left had coordinate system.
                    if ((inputPort.Equals(boundingPort)) && (wAxis * uAxis == vAxis))
                        wAxis.Length = -1.0;
                }
            }

            // ------------------------------------------------------
            // If flange is penetrated, swap u and v, and negate v
            // ------------------------------------------------------
            if (sketchPlane == null && !IsWebPenetrating())
            {
                if (memberAxisPort != null)
                {
                    if (memberAxisPort.AxisPortType == MemberAxisPortType.Start || memberAxisPort.AxisPortType == MemberAxisPortType.End ||
                        (inputPort.Equals(boundedPort) && memberAxisPort.AxisPortType == MemberAxisPortType.Along))
                    {
                        Vector tempVec = new Vector(uAxis);
                        uAxis = new Vector(vAxis);
                        vAxis = new Vector(tempVec);
                            vAxis.Length = -1.0;
                    }
                }
                else if (oTopoPort != null)
                {
                    if (oTopoPort.ContextId == ContextTypes.Base || oTopoPort.ContextId == ContextTypes.Offset)
                    {
                        Vector tempVec = new Vector(uAxis);
                        uAxis = new Vector(vAxis);
                        vAxis = new Vector(tempVec);                        

                    vAxis.Length = -1.0;
                }
            }
        }
        }
        /// <summary>
        /// Determines whether or not the bounding axis is in the same direction as the sketching plane normal
        /// </summary>
        /// <param name="boundingPort">The bounding port.</param>
        /// <param name="boundedPort">The bounded port.</param>
        /// <returns>
        /// 	<c>true</c> if the bounding axis is in the same direction as the sketching plane normal; otherwise, <c>false</c>.
        /// </returns>
        /// <exception cref="ArgumentNullException">The passed port is null.</exception>
        internal bool IsPenetrationWithSketchingPlaneNormal()
        {
            // Don't confuse this with the Mirror property on MemberPart, which does not help determine the member is mirrored
            // in the sense of the endcut sketching plane.
            //
            // Also, this is used for more than determining if left and right should be flipped.  It is also relevant when determining
            // the orientation quadrant of the bounding section.  The orientation quadrant is relative to it primary orientation
            // (v-direction of the cross section), whereas left and right are relative to the secondary orientation.  Nonetheless, due to the
            // algorithm chosen, it is a factor.
            Vector u, v, w, U, V, W;

            GetEndCutLocalCoordinateSystem(boundingPort, out u, out v, out w);
            GetEndCutLocalCoordinateSystem(boundedPort, out U, out V, out W);

            Vector SketchPlaneNormal;
//            if (IsWebPenetrating())
//            {
                SketchPlaneNormal = V * W;
//            }
//            else
//            {
//                SketchPlaneNormal = V;  //Plus or Minus V?
//            }
            double dotSw = System.Math.Round(SketchPlaneNormal % w, 4);
            MemberPart boundedMember = boundedPort.Connectable as MemberPart;
            if (boundedMember != null)
            {
                if (IsMemberATubeType(boundedMember) == true)
                {
                    StiffenerPartBase StiffenerObj = boundingPort.Connectable as StiffenerPartBase;
                    if (StiffenerObj != null)
                    {
                        dotSw = -1.0; //Always return false when tube-member is bounded to stiffener
                    }
                }
            }

            return (dotSw  > 0.0);
        }

        /// <summary>
        /// Gets the orientation quadrant.
        /// </summary>
        /// <param name="boundingMemberPort">The bounding member port.</param>
        /// <param name="boundedMemberPort">The bounded member port.</param>
        /// <returns>The quadrandt (1, 2, 3, or 4)</returns>
        /// <exception cref="S3DEndCutMappingException">Unknown configuration for quadrant calculation.</exception>
        internal int GetOrientationQuadrant()
        {
            int quadrant = 1;

            // Get the local coordinate systems adjust for endcut calculation
            Vector u, v, w, U, V, W;

            GetEndCutLocalCoordinateSystem(boundingPort, out u, out v, out w);
            GetEndCutLocalCoordinateSystem(boundedPort, out U, out V, out W);

            // Get cosine 45deg for comparison purposes. Round it to 4 decimals.
            double cos45 = System.Math.Round(Math3d.CosDeg(45), 4);

            //Check if earlier logic for quadrant mapping can be used: this seem to work only
            //when W and w vectors are at angle > 45 degree, means their dot product is 
            // less than cos(45)
            double dot_Ww = System.Math.Round(W % w, 4);
            if (System.Math.Abs(dot_Ww) < cos45)
            {
                //Follow earlier logic
            }
            else
            {
                //Re-write W (bounded vector): use existing W vector's projection on to the plane 
                // perpendicular to the w(bounding axis): this projected vector is new W
                Vector cross_Ww = W * w;
                cross_Ww.Length = 1.0;

                Vector New_W = cross_Ww * w;
                New_W.Length = 1.0;
                //Ensure that New_W maintains same orientation as that of W
                if (System.Math.Round(W % New_W, 4) < 0.0)
                {
                    New_W.Length = -1.0;
                }
                W = New_W;
            }

            // Get the 2 dot products which will determine initial quadrant selection. Round it to 4 decimals.
            double dot_Wu = System.Math.Round(W % u, 4);
            double dot_Wv = System.Math.Round(W % v, 4);

            // Determine if the bounding axis and bounded web right are in the same direction
            bool penetrationWithNormal = IsPenetrationWithSketchingPlaneNormal();

            if (dot_Wv > cos45)
            {
                quadrant = 2;
            }
            else if (dot_Wv < -cos45)
            {
                quadrant = 4;
            }
            else if (dot_Wu >= cos45)
            {
                if (penetrationWithNormal)
                {
                    quadrant = 1;
                }
                else
                {
                    quadrant = 3;
                }
            }
            else if (dot_Wu <= -cos45)
            {
                if (penetrationWithNormal)
                {
                    quadrant = 3;
                }
                else
                {
                    quadrant = 1;
                }
            }
            else
            {
                throw new EdgeMappingException("Unknown configuration for quadrant calculation, W.u = " + dot_Wu + " W.v = " + dot_Wv);
            }

            return quadrant;
        }

        /// <summary>
        /// Determines whether the member is of Tube type.
        /// </summary>
        /// <param name="oMemberPart">The member.</param>
        /// <returns>
        /// 	<c>true</c> if is tube type, otherwise, <c>false</c>.
        /// </returns>
        internal bool IsMemberATubeType(MemberPart oMemberPart)
        {
            bool bIsTube = false;

            if (oMemberPart != null)
            {
                string sectionType = oMemberPart.SectionType;
                
                switch (sectionType)
                {
                    case "HSSC":
                    case "CS":
                    case "PIPE":
                        bIsTube = true;
                        break;
                }
            }
            return bIsTube;
        }
        
        
        /// <summary>
        /// Determines whether the bounding member axis is penetrating the web of specified bounded member.
        /// </summary>
        /// <param name="boundingMember">The bounding member.</param>
        /// <param name="boundedMember">The bounded member.</param>
        /// <returns>
        /// 	<c>true</c> if web penetrating, otherwise, <c>false</c>.
        /// </returns>
        internal bool IsWebPenetrating()
        {
            MemberPart boundingMember = boundingPort.Connectable as MemberPart;
            MemberPart boundedMember = boundedPort.Connectable as MemberPart;
            StiffenerPartBase boundingStiffener = boundingPort.Connectable as StiffenerPartBase;
            StiffenerPartBase boundedStiffener = boundedPort.Connectable as StiffenerPartBase;

            // if Bounded is Tube and Bounding is Not Tube, consider as WebPenetrating
            if ((IsMemberATubeType(boundedMember) == true) && (IsMemberATubeType(boundingMember)== false))
            {
                return true;
            }

            bool isWebPenetrating = true; //by default web is penetrating

            //*** Noticed that for curved members using XAxis, YAxis and ZAxis 
            // seem to give incorrect result (these vectors are defined at member
            // start position: where as for curved member we are are intrested in
            // calculating dot product of vectors at bounding location ***
            // for now kept earlier code commented-out

            //    x      w   v                          x        w    v
            //    ^      ^   ^                          ^        ^    ^
            //    |      |   |                          |        |    |
            //     |      |  |                          |         |   |
            //      |      | |                          |          |  |
            //      penetrates web                     penetrates flange
            //

            //Vector w = boundedMember.YAxis; //w is bounded Y axis
            //w.Length = 1.0;
            //Vector v = boundedMember.ZAxis; // v is bounded Z axis
            //v.Length = 1.0;
            //Vector x = boundingMember.XAxis;
            //x.Length = 1.0;

            //// make sure all are unit vectors
            //double dot_wx = System.Math.Round(System.Math.Abs(x % w), 4);
            //double dot_vx = System.Math.Round(System.Math.Abs(x % v), 4);

            //if v-vector is more alligned with the bounding x-axis, must be penetrating flange
            Vector w, U, V;

            //Step 1: compute vectors for bounding
            w = new Vector();
            if (boundingMember != null)
            {
            Matrix4X4 oboundingMbrMatrix = boundingMember.GetMatrixAtPosition(this.boundedLocation);

            w.Set(oboundingMbrMatrix.GetIndexValue(0), oboundingMbrMatrix.GetIndexValue(1), oboundingMbrMatrix.GetIndexValue(2));
            }
            else if (boundingStiffener != null)
            {
                Matrix4X4 oboundingStfMatrix = boundingStiffener.GetMatrixAtPosition(this.boundedLocation);
                w.Set(oboundingStfMatrix.GetIndexValue(8), oboundingStfMatrix.GetIndexValue(9), oboundingStfMatrix.GetIndexValue(10));
                //w = new Vector(boundingStiffener.XAxis);
            }
            else
            {
                //Unknown case. might be an error or need to handle ???
            }

            //Step 2: compute vectors for bounded
            U = new Vector();
            V = new Vector();
            if (boundedMember != null)
            {
                Matrix4X4 oboundedMbrMatrix = boundedMember.GetMatrixAtPosition(this.boundedLocation);
                U.Set(oboundedMbrMatrix.GetIndexValue(4), oboundedMbrMatrix.GetIndexValue(5), oboundedMbrMatrix.GetIndexValue(6));
            V.Set(oboundedMbrMatrix.GetIndexValue(8), oboundedMbrMatrix.GetIndexValue(9), oboundedMbrMatrix.GetIndexValue(10));
            }
            else if (boundedStiffener != null)
            {
                Matrix4X4 oboundedStfMatrix = boundedStiffener.GetMatrixAtPosition(this.boundedLocation);
                U.Set(oboundedStfMatrix.GetIndexValue(0), oboundedStfMatrix.GetIndexValue(1), oboundedStfMatrix.GetIndexValue(2));
                V.Set(oboundedStfMatrix.GetIndexValue(4), oboundedStfMatrix.GetIndexValue(5), oboundedStfMatrix.GetIndexValue(6));
                //U = new Vector(boundedStiffener.YAxis);
                //V = new Vector(boundedStiffener.ZAxis);
            }
            else
            {
                //Unknown case. might be an error or need to handle ???
            }

            w.Length = 1.0;
            U.Length = 1.0;
            V.Length = 1.0;

            double dot_wU = System.Math.Round(System.Math.Abs(U % w), 4);
            double dot_wV = System.Math.Round(System.Math.Abs(V % w), 4);

            //If wV > wU
            if (System.Math.Abs(dot_wV) > System.Math.Abs(dot_wU))
            {
                isWebPenetrating = false;
            }

            return isWebPenetrating;
        }

        /// <summary>
        /// Gets the sketching plane to be used when the bounded object is tubular
        /// </summary>
        /// <param name="boundingMemberPort">The bounding member port.</param>
        /// <param name="boundedMemberPort">The bounded member port.</param>
        internal Plane3d GetSketchPlaneForTube()
        {
            // -----------------------------------------------------------------
            // Return a null object if the bounded port is not a member end port
            // -----------------------------------------------------------------
            MemberPartAxisPort boundedAxisPort = boundedPort as MemberPartAxisPort;
            if (boundedAxisPort == null)
            {
                return null;
            }

            if (boundedAxisPort.AxisPortType != MemberAxisPortType.Start && boundedAxisPort.AxisPortType != MemberAxisPortType.End)
            {
                 return null;
            }

            MemberPart boundedPart = boundedPort.Connectable as MemberPart;

            // -----------------------
            // Get the bouned location
            // -----------------------
            Position boundedStart;
            Position boundedEnd;
            Position boundedPos;

            //...boundedPart.GetEndPoints(out boundedStart, out boundedEnd);
            // GetEndPoints method call raises an exception for curved member part,
            // so use an alternative approach
            IPoint boundedStPos = boundedPart.GetPointAtEnd(MemberAxisEnd.Start);
            // Initialize the variables.
            double x = 0.0, y = 0.0, z = 0.0;
            x = boundedStPos.X;
            y = boundedStPos.Y;
            z = boundedStPos.Z;
            boundedStart = new Position();
            boundedStart.Set(x, y, z);

            IPoint boundedEnPos = boundedPart.GetPointAtEnd(MemberAxisEnd.End);
            // Initialize the variables.
            x = boundedEnPos.X;
            y = boundedEnPos.Y;
            z = boundedEnPos.Z;
            boundedEnd = new Position();
            boundedEnd.Set(x, y, z);

            if (boundedAxisPort.AxisPortType == MemberAxisPortType.Start)
            {
                 boundedPos = boundedStart;
            }
            else
            {
                boundedPos = boundedEnd;
            }

            // ----------------------------------------------------------
            // Get the bounded orientation matrix at the bounded location
            // ----------------------------------------------------------
            Matrix4X4 boundedMatrix = boundedPart.GetMatrixAtPosition(boundedPos);

            // --------------------------------------------------------------------------------------------------
            // Sketching plane U is the axis direction at the bounded location, adjusted to point toward the part
            // --------------------------------------------------------------------------------------------------
            Vector sketchU = new Vector(boundedMatrix.GetIndexValue(0), boundedMatrix.GetIndexValue(1), boundedMatrix.GetIndexValue(2));

            if (boundedAxisPort.AxisPortType == MemberAxisPortType.End)
            {
                sketchU.Length = -1.0;
            }

            // ---------------------------
            // If bounded by a member axis 
            // ---------------------------
            Vector boundingAxis;
            Position sketchRoot;
            MemberPart boundingMember = boundingPort.Connectable as MemberPart;
            PlatePart boundingPlate = boundingPort.Connectable as PlatePart;
            StiffenerPartBase boundingStiffener = boundingPort.Connectable as StiffenerPartBase;

            if ((boundingMember != null) || (boundingStiffener != null))
            {
                if (boundingMember != null) //bounding is a standard member
                {
                    MemberPartAxisPort boundingAxisPort = boundingPort as MemberPartAxisPort;
                    if (boundingAxisPort != null)
                    {
                        if (boundingAxisPort.AxisPortType != MemberAxisPortType.Along)
                        {
                            return null;
                        }
                    }
                    // for Generic Connections, the bounding Port may not be Axis Along,
                    // hence check whether the bounding Port type is atleast FACE
                    else
                    {
                        if (boundingPort.PortType != Ingr.SP3D.Common.Middle.PortType.Face)
                        {
                            return null;
                        }
                    }
                    // -------------------------------------------------------------
                    // The bounding axis is the member axis at the bounded location
                    // -------------------------------------------------------------
                    Matrix4X4 boundingMatrix = boundingMember.GetMatrixAtPosition(boundedPos);

                    boundingAxis = new Vector(boundingMatrix.GetIndexValue(0), boundingMatrix.GetIndexValue(1), boundingMatrix.GetIndexValue(2));
                }
                else //bounding is a Stiffener
                {
                    // for Generic Connection: check whether the bounding Port type is FACE
                    if (boundingPort.PortType != Ingr.SP3D.Common.Middle.PortType.Face)
                    {
                        return null;
                    }
                    //Bounding is a stiffener
                    // -------------------------------------------------------------
                    // The bounding axis is the member axis at the bounded location
                    // -------------------------------------------------------------
                    Matrix4X4 boundingMatrix = boundingStiffener.GetMatrixAtPosition(boundedPos);
                    MemberPartAxisPort boundingAxisPort = boundingPort as MemberPartAxisPort;
                    boundingAxis = new Vector(boundingMatrix.GetIndexValue(8), boundingMatrix.GetIndexValue(9), boundingMatrix.GetIndexValue(10));
                }
                // -------------------------------------------------------
                // The sketching plane root is at load point 15 of bounded
                // -------------------------------------------------------
                Vector boundedU = new Vector(boundedMatrix.GetIndexValue(4), boundedMatrix.GetIndexValue(5), boundedMatrix.GetIndexValue(6));
                Vector boundedV = new Vector(boundedMatrix.GetIndexValue(8), boundedMatrix.GetIndexValue(9), boundedMatrix.GetIndexValue(10));

                double u15 = 0.0;
                double v15 = 0.0;
                int CardinalPointofBoundedPart = 0;

                Ingr.SP3D.Structure.Middle.Services.CrossSectionServices MemberPartHelper = new Ingr.SP3D.Structure.Middle.Services.CrossSectionServices();
                ProfilePart oProfile = boundedPart;

                CardinalPointofBoundedPart = boundedPart.CardinalPoint;

                MemberPartHelper.GetCardinalPointDelta(oProfile, CardinalPointofBoundedPart, 15, out u15, out v15);

                boundedU.Length = -u15;
                boundedV.Length = v15;
                sketchRoot = new Position(boundedMatrix.GetIndexValue(12) + boundedU.X + boundedV.X, boundedMatrix.GetIndexValue(13) + boundedU.Y + boundedV.Y, boundedMatrix.GetIndexValue(14) + boundedU.Z + boundedV.Z);

            }
            // --------------------------
            // If bounded by a plate edge
            // --------------------------
            else if (boundingPlate != null)
            {
                // --------------------------
                // Exit if not a lateral port
                // --------------------------
                TopologyPort topoPort = boundingPort as TopologyPort;
                if ((topoPort.ContextId & ContextTypes.Lateral) != ContextTypes.Lateral)
                {
                    return null;
                }

                // ----------------------------------------------------------------------------------------------
                // Bounding axis (as if the plate were a member) is the plate normal nearest the bounded location
                // ----------------------------------------------------------------------------------------------
                ISurface basePort = boundingPlate.GetPort(TopologyGeometryType.Face, -1, -1, ContextTypes.Base, -1, true) as ISurface;
                if (basePort == null)
                {
                    throw new ArgumentNullException("bounding plate base port");
                }

                Position ptOnPort = basePort.ProjectPoint(boundedPos);
                boundingAxis = basePort.OutwardNormalAtPoint(ptOnPort);

                // ---------------------------------------------
                // sketchRoot is centered in the plate thickness
                // ---------------------------------------------
                Vector offsetVector = new Vector(boundingAxis);
                offsetVector.Length = -boundingPlate.Thickness;
                sketchRoot = ptOnPort.Offset(offsetVector);
            }
            else
            {
                return null; //Unknown bounding type
            }

            // --------------------------------------------------------------
            // Sketch V is the cross product of the bounding axis and sketchU
            // --------------------------------------------------------------
            Vector sketchV = boundingAxis.Cross(sketchU);

             // ----------------------------------------------------
             // Sketch N is the cross product of sketchU and sketchV
             // ----------------------------------------------------
             Vector sketchN = sketchU.Cross(sketchV);

             sketchN.Length = 1.0;
             sketchU.Length = 1.0;

             Plane3d sketchPlane = new Plane3d(sketchRoot, sketchN);
             sketchPlane.UDirection = sketchU;
             return sketchPlane;
        }

        /// <summary>
        /// Determines if the current quadrant is good enough mapping 
        /// for the bounded to bounding orientation
        /// for e.g. there are cases where mapping depends not only on
        /// bounded/bounding U,V & W vectors but instead but even could 
        /// be on the orentation specific
        /// Below method determines such cases and returns approprate
        /// output parameters
        /// </summary>
        /// <param name="currentQuadrant">The quadrant before remapping.</param>
        /// <param name="sectionAlias">Section Alias before remapping.</param>
        /// <param name="edges">Mapped edges collection before remapping.</param>
        /// <returns>
        /// 	<c>newQuadrant</c> new quadrant if remap is needed</c>.
        /// 	<c>bHasToReMap</c> boolean flag to control remapping</c>.
        /// </returns>
        internal void ReMapIfNeeded (int currentQuadrant, int sectionAlias, Dictionary<int,int> edges, out int newQuadrant, out bool bHasToReMap)
        {
            //set default values
            newQuadrant = currentQuadrant;
            bHasToReMap = false;

            //currently only handled for Web penetration
            if (!IsWebPenetrating())
            {
                return;
            }

            //currently supports only members 
            MemberPart oBoundedMember = boundedPort.Connectable as MemberPart;
            MemberPart oBoundingMember = boundingPort.Connectable as MemberPart;

            if ((oBoundedMember == null) || (oBoundingMember == null))
            {
                // return -- NO REMAPPING IS NEEDED
                //currently handled only for Members
                return;
            }


            double dAngularTolerance = 15;  // 15 degrees as tolerance level
            bool bValidate = false;
            int nAdjacentSectionFaceType = 0;
            int nAdjacentQuadrant = currentQuadrant;
            
            Plane3d oSketchPlane = null;

            bValidate = ValidateConditionsForRemapping(currentQuadrant, sectionAlias, dAngularTolerance, out nAdjacentQuadrant, out nAdjacentSectionFaceType, out oSketchPlane);                    
 

            if ((bValidate) && (nAdjacentQuadrant != currentQuadrant))
            {
                //everything is handled with respect to Web penetrated cases
                //for other cases(for e.g Flange Penetrated cases) need to enhance the method
                Surface3d oBoundedTopPort = oBoundedMember.GetExtendedLateralSurface((int)SectionFaceType.Top, 2);
                Surface3d oBoundedBtmPort = oBoundedMember.GetExtendedLateralSurface((int)SectionFaceType.Bottom, 2);

                if ((oBoundedTopPort == null) || (oBoundedBtmPort == null))
                {
                    //unusual case(for e.g for Tube as bounded)!!
                    //Neeed to handle as per the requirement
                    return;
                }

                if (oSketchPlane == null)
                {
                    //we need to have sketch plane filled
                    //for doing geometry calc below.
                    //if Sketch plane is NULL 
                    //it is error case but as of now
                    //exit safely
                    //WE HAVE TO HAVE SKETCH PLANE TO GO AHEAD..!!!
                    return;
                }

                int nMappedWRPort;
                int nMappedAdjacentPort;

                edges.TryGetValue((int)SectionFaceType.Web_Right, out nMappedWRPort);
                edges.TryGetValue(nAdjacentSectionFaceType, out nMappedAdjacentPort);

                TopologyPort oMappedWRPort = oBoundingMember.GetPort(TopologyGeometryType.Face, 0, nMappedWRPort, ContextTypes.Lateral, nMappedWRPort, false);
                TopologyPort oMappedAdjacentPort = oBoundingMember.GetPort(TopologyGeometryType.Face, 0, nMappedAdjacentPort, ContextTypes.Lateral, nMappedAdjacentPort, false);

                GeometryIntersectionType eIntersectionType;
                Collection<ICurve> colCurves;

                oBoundedTopPort.Intersect((ISurface)oSketchPlane, out colCurves, out eIntersectionType);
                ICurve oBdedTopEdge = null;

                if (colCurves != null)
                {
                    if (colCurves.Count > 0)
                    {   
                        // there is supposed to be only one count
                        oBdedTopEdge = colCurves[0];
                    }
                }
                else
                {
                    return;
                }

                colCurves.Clear(); // clear the collection before filling some other collection
                oBoundedBtmPort.Intersect((ISurface)oSketchPlane, out colCurves, out eIntersectionType);
                ICurve oBdedBtmEdge = null;

                if (colCurves != null)
                {
                    if (colCurves.Count > 0)
                    {
                        oBdedBtmEdge = colCurves[0];
                    }
                }
                else
                {
                    return;
                }

                colCurves.Clear(); // clear the collection before filling some other collection
                oMappedWRPort.Intersect((ISurface)oSketchPlane, out colCurves, out eIntersectionType);
                ICurve oBdingMappedWREdge = null;

                if (colCurves != null)
                {
                    if (colCurves.Count > 0)
                    {
                        oBdingMappedWREdge = colCurves[0];
                    }
                }
                else
                {
                    return;
                }

                colCurves.Clear(); // clear the collection before filling some other collection
                oMappedAdjacentPort.Intersect((ISurface)oSketchPlane, out colCurves, out eIntersectionType);
                ICurve oBdingMappedAdjacentEdge = null;

                if (colCurves != null)
                {
                    if (colCurves.Count > 0)
                    {
                        oBdingMappedAdjacentEdge = colCurves[0];
                    }
                }
                else
                {
                    return;
                }

                colCurves.Clear(); // clear the collection before filling some other collection

                Collection<Position> colTopIntersectPos;
                Collection<Position> colTopOverlapPos;
                Collection<Position> colBtmIntersectPos;
                Collection<Position> colBtmOverlapPos;

                oBdingMappedAdjacentEdge.Intersect(oBdedTopEdge, out colTopIntersectPos, out colTopOverlapPos, out eIntersectionType);
                oBdingMappedAdjacentEdge.Intersect(oBdedBtmEdge, out colBtmIntersectPos, out colBtmOverlapPos, out eIntersectionType);

                Matrix4X4 oMatrix = new Matrix4X4();
                oMatrix.SetIndexValue(0, oSketchPlane.UDirection.X);
                oMatrix.SetIndexValue(1, oSketchPlane.UDirection.Y);
                oMatrix.SetIndexValue(2, oSketchPlane.UDirection.Z);
                oMatrix.SetIndexValue(4, oSketchPlane.VDirection.X);
                oMatrix.SetIndexValue(5, oSketchPlane.VDirection.Y);
                oMatrix.SetIndexValue(6, oSketchPlane.VDirection.Z);
                oMatrix.SetIndexValue(8, oSketchPlane.Normal.X);
                oMatrix.SetIndexValue(9, oSketchPlane.Normal.Y);
                oMatrix.SetIndexValue(10, oSketchPlane.Normal.Z);
                oMatrix.SetIndexValue(12, oSketchPlane.RootPoint.X);
                oMatrix.SetIndexValue(13, oSketchPlane.RootPoint.Y);
                oMatrix.SetIndexValue(14, oSketchPlane.RootPoint.Z);
                oMatrix.Invert();

                if ((colTopIntersectPos != null) && (colBtmIntersectPos != null))
                {
                    if ((colTopIntersectPos.Count > 0) && (colBtmIntersectPos.Count > 0))
                    {
                        bHasToReMap = true;
                        newQuadrant = nAdjacentQuadrant;
                    }
                }
                else if ((colTopIntersectPos != null) && (colBtmIntersectPos == null) || 
                    (colTopIntersectPos == null) && (colBtmIntersectPos != null))
                {
                    Position posStartPoint;
                    Position posEndPoint;
                    oBdingMappedAdjacentEdge.EndPoints(out posStartPoint, out posEndPoint);

                    Position pos2dStartPoint = oMatrix.Transform(posStartPoint);
                    Position pos2dEndPoint = oMatrix.Transform(posEndPoint);

                    Position posNearMost = new Position();

                    if ((pos2dStartPoint.X) >= (pos2dEndPoint.X))
                    {
                        posNearMost = posStartPoint;
                    }
                    else
                    {
                        posNearMost = posEndPoint;
                    }

                    Position posOnTop = oBdedTopEdge.ProjectPoint(posNearMost, oSketchPlane.VDirection);
                    Position posOnBtm = oBdedBtmEdge.ProjectPoint(posNearMost, oSketchPlane.VDirection);


                    if ((posOnTop == null) || (posOnBtm == null))
                    {
                        //!!! unknown case, Need to handle as per the cases
                        //we never expect the code to fail here
                        //if fails need to handle appropriately, currently we
                        //safely exit
                        return;
                    }

                    if (((posNearMost.DistanceToPoint(posOnTop)) + (posNearMost.DistanceToPoint(posOnBtm))) > posOnTop.DistanceToPoint(posOnBtm))
                    {

                        bHasToReMap = true;
                        newQuadrant = nAdjacentQuadrant;

                    }
                    else
                    {
                        if (((colTopIntersectPos != null) && (posNearMost.DistanceToPoint(posOnTop) > posNearMost.DistanceToPoint(posOnBtm))) ||
                            ((colBtmIntersectPos != null) && (posNearMost.DistanceToPoint(posOnBtm) > posNearMost.DistanceToPoint(posOnTop))))
                        {
                            bHasToReMap = true;
                            newQuadrant = nAdjacentQuadrant;
                        }
                        else
                        {
                            //no need to remapp
                            return;
                        }
                    }


                }
                else
                {
                    Position posStartPoint;
                    Position posEndPoint;
                    oBdingMappedAdjacentEdge.EndPoints(out posStartPoint, out posEndPoint);

                    Position pos2dStartPoint = oMatrix.Transform(posStartPoint);
                    Position pos2dEndPoint = oMatrix.Transform(posEndPoint);

                    Position posNearMost = new Position();

                    if ((pos2dStartPoint.X) >= (pos2dEndPoint.X))
                    {
                        posNearMost = posStartPoint;
                    }
                    else
                    {
                        posNearMost = posEndPoint;
                    }

                    Position posOnTop = oBdedTopEdge.ProjectPoint(posNearMost, oSketchPlane.VDirection);
                    Position posOnBtm = oBdedBtmEdge.ProjectPoint(posNearMost, oSketchPlane.VDirection);

                    if ((posOnTop == null) || (posOnBtm == null))
                    {
                        //!!! unknown case, Need to handle as per the cases
                        //we never expect the code to fail here
                        //if fails need to handle appropriately, currently we
                        //safely exit
                        return;
                    }

                    if (nAdjacentSectionFaceType == (int)SectionFaceType.Top)
                    {
                        if (posNearMost.DistanceToPoint(posOnTop) > posNearMost.DistanceToPoint(posOnBtm))
                        {
                            bHasToReMap = true;
                            newQuadrant = nAdjacentQuadrant;
                        }
                    }
                    else if (nAdjacentSectionFaceType == (int)SectionFaceType.Bottom)
                    {
                        if (posNearMost.DistanceToPoint(posOnBtm) > posNearMost.DistanceToPoint(posOnTop))
                        {
                            bHasToReMap = true;
                            newQuadrant = nAdjacentQuadrant;
                        }
                    }
                    else
                    {
                        // other cases need to consider as per the need
                    }
                }
            }

            return;
               
        }
       
  
        /// <summary>
        /// This is a helper method to validate the conditions for 
        /// remapping.
        /// Remapping stuff is only needed/done for 
        ///    1) particular angular orienatations within tolerance level
        ///    2) Particular Section Aliases for specific Quadrants
        /// if Remapping found true then returns
        ///     ----quadrant adjacent ot the current quadrant
        ///     ----Xid of AdjacentPort of Mapped Web right Port and 
        ///     ----End Cut Sketch Plane needed for Geom Calc
        /// </summary>
        internal bool ValidateConditionsForRemapping(int currentQuadrant, int sectionAlias, double dAngularTolerance, out int nAdjacentQuadrant, 
                                                    out int nAdjacentMappedSectionFaceType, out Plane3d oSketchPlane)
        {
            
            //set default values
            nAdjacentQuadrant = currentQuadrant;
            nAdjacentMappedSectionFaceType = (int)SectionFaceType.Unknown;
            oSketchPlane = null;
            bool bValidated = false;

            // Get the local coordinate systems adjust for endcut calculation
            Vector u, v, w, U, V, W;
            
            GetEndCutLocalCoordinateSystem(boundingPort, out u, out v, out w);
            GetEndCutLocalCoordinateSystem(boundedPort, out U, out V, out W);

            Vector Projetced_W = W; //default value
            double dot_Ww = System.Math.Round(W % w, 4);

            if ((System.Math.Abs(dot_Ww)).EqualTo(0)== false)
            {
                Vector cross_Ww = W * w;
                cross_Ww.Length = 1.0;

                Projetced_W = cross_Ww * w;
                Projetced_W.Length = 1.0;


                //Ensure that Projetced_W maintains same orientation as that of W
                if (System.Math.Round(W % Projetced_W, 4) < 0.0)
                {
                    Projetced_W.Length = -1.0;
                }
            }

            // Get cosine 45deg for comparison purposes. Round it to 4 decimals.
            double cos45 = System.Math.Round(Math3d.CosDeg(45), 4);
            double cosTheta = System.Math.Round(Math3d.CosDeg(45 - dAngularTolerance), 4);

            double dot_Wu = System.Math.Round(Projetced_W % u, 4);
            double dot_Wv = System.Math.Round(Projetced_W % v, 4);

            if (currentQuadrant == 1)
            {
                if ((dot_Wu <= -cos45) && (dot_Wu >= -cosTheta) && (dot_Wv > 0) ||   //non flipped condition
                    (dot_Wu >= cos45) && (dot_Wu <= cosTheta) && (dot_Wv > 0))       //Flipped condition
                {
                    if ((sectionAlias == (int)SectionAliasType.WebTopAndBottomFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebBottomFlangeRight) ||
                        (sectionAlias == (int)SectionAliasType.WebBottomFlangeLeft) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomRightFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomLeftFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebBottomFlange) ||
                        (sectionAlias == (int)SectionAliasType.TwoWebsTwoFlanges) ||
                        (sectionAlias == (int)SectionAliasType.FlangeLeftAndRightTopWebs))
                    {
                        nAdjacentQuadrant = 2;
                        nAdjacentMappedSectionFaceType = (int)SectionFaceType.Bottom;
                        bValidated = true;
                    }
                    else
                    {
                        //remapping not required
                        bValidated = false;
                    }
                }
                else if ((dot_Wu <= -cos45) && (dot_Wu >= -cosTheta) && (dot_Wv < 0) || //non flipped condition
                        (dot_Wu >= cos45) && (dot_Wu <= cosTheta) && (dot_Wv < 0))      //Flipped condition
                {
                    if ((sectionAlias == (int)SectionAliasType.WebTopAndBottomFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebTopFlange) ||
                        (sectionAlias == (int)SectionAliasType.WebTopFlangeLeft) ||
                        (sectionAlias == (int)SectionAliasType.WebTopFlangeRight) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomRightFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomLeftFlanges) ||
                        (sectionAlias == (int)SectionAliasType.TwoWebsTwoFlanges) ||
                        (sectionAlias == (int)SectionAliasType.FlangeLeftAndRightBottomWebs))
                    {
                        nAdjacentQuadrant = 4;
                        nAdjacentMappedSectionFaceType = (int)SectionFaceType.Top;
                        bValidated = true;
                    }
                    else
                    {
                        //remapping not required
                        bValidated = false;
                    }
                }
                else
                {
                    //remapping not required since
                    //doesnt lie in the required Tolerance
                    bValidated = false;
                }
            }
            else if (currentQuadrant == 3)
            {
                if ((dot_Wu >= cos45) && (dot_Wu <= cosTheta) && (dot_Wv < 0) ||   //non flipped condition
                    (dot_Wu <= -cos45) && (dot_Wu >= -cosTheta) && (dot_Wv < 0))   //Flipped Conditions
                {
                    if ((sectionAlias == (int)SectionAliasType.WebTopAndBottomFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebBottomFlangeRight) ||
                        (sectionAlias == (int)SectionAliasType.WebBottomFlangeLeft) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomRightFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomLeftFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebBottomFlange) ||
                        (sectionAlias == (int)SectionAliasType.TwoWebsTwoFlanges) ||
                        (sectionAlias == (int)SectionAliasType.FlangeLeftAndRightTopWebs))
                    {
                        nAdjacentQuadrant = 4;
                        nAdjacentMappedSectionFaceType = (int)SectionFaceType.Bottom;
                        bValidated = true;
                    }
                    else
                    {
                        //remapping not required
                        bValidated = false;
                    }                    
                }
                else if ((dot_Wu >= cos45) && (dot_Wu <= cosTheta) && (dot_Wv > 0) ||   //non flipped condition
                        (dot_Wu <= -cos45) && (dot_Wu >= -cosTheta) && (dot_Wv > 0))    ///Flipped Conditions
                {
                    if ((sectionAlias == (int)SectionAliasType.WebTopAndBottomFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebTopFlange) ||
                        (sectionAlias == (int)SectionAliasType.WebTopFlangeLeft) ||
                        (sectionAlias == (int)SectionAliasType.WebTopFlangeRight) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomRightFlanges) ||
                        (sectionAlias == (int)SectionAliasType.WebTopAndBottomLeftFlanges) ||
                        (sectionAlias == (int)SectionAliasType.TwoWebsTwoFlanges) ||
                        (sectionAlias == (int)SectionAliasType.FlangeLeftAndRightBottomWebs))
                    {
                        nAdjacentQuadrant = 2;
                        nAdjacentMappedSectionFaceType = (int)SectionFaceType.Top;
                        bValidated = true;
                    }
                    else
                    {
                        //remapping not required
                        bValidated = false;
                    }
                }
                else
                {
                    //remapping not required since
                    //doesnt lie in the required Tolerance
                    bValidated = false;
                }
            }
            else if (currentQuadrant == 2)
            {
                // not yet implemented
                bValidated = false;
            }
            else if (currentQuadrant == 4)
            {
                // not yet implemented
                bValidated = false;
            }

            //Evaluate Sketch Plane
            if (bValidated)
            {
                MemberPart oBoundedMember = boundedPort.Connectable as MemberPart;
                

                if (oBoundedMember != null) 
                {
                    ISurface oWebLeftSurface = (ISurface)oBoundedMember.GetPort(TopologyGeometryType.Face, 0, (int)SectionFaceType.Web_Left, ContextTypes.Lateral, (int)SectionFaceType.Web_Left, false);

                    if (oWebLeftSurface != null)
                    {

                        Position oSketchPlaneRootPosition;
                        Position posOnWebLeft = null;
                        Position posSrc;
                        Position posIn;

                        double dMinDist;
                        double dWebThickness;

                        Point3d oBoundedLoc = new Point3d(boundedLocation);
                        ISurface oWebRightSurface = (ISurface)oBoundedMember.GetPort(TopologyGeometryType.Face, 0, (int)SectionFaceType.Web_Right, ContextTypes.Lateral, (int)SectionFaceType.Web_Right, false);

                        if (oWebRightSurface == null)
                        {
                            //peculiar cases
                            //if code stops here we have to find out
                            //some other way to detrermin Flange Thickness
                            //Currently safely exit
                            return bValidated;
                        }

                        oWebRightSurface.DistanceBetween(oWebLeftSurface, out dWebThickness, out posSrc, out posIn);

                        Vector oSketchPlaneUDir = new Vector(-W.X, -W.Y, -W.Z);
                        oWebLeftSurface.DistanceBetween(oBoundedLoc, out dMinDist, out posOnWebLeft);
                        Vector oSketchOffsetDir = oWebLeftSurface.OutwardNormalAtPoint(posOnWebLeft);
                        oSketchOffsetDir.Length = -(dWebThickness / 2);

                        oSketchPlaneRootPosition = posOnWebLeft.Offset(oSketchOffsetDir);
                        Plane3d oInternalSketchPlane = new Plane3d(oSketchPlaneRootPosition, oSketchPlaneUDir.Cross(V));
                        oInternalSketchPlane.UDirection = oSketchPlaneUDir;
                        oSketchPlane = oInternalSketchPlane;

                    }
                }
                else
                {
                }
            }
            return bValidated;

        }
        
    }
}
