using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Exceptions;
using System.Collections;

using Ingr.SP3D.Manufacturing.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    public class MarginRule:MarginRuleBase
    {

        /// <summary>
        /// Gets the allowed directions to apply fabrication margin.
        /// </summary>
        /// <param name="marginInformation">Input margin information.</param>
        /// <returns>Returns the list of allowed directions.</returns>
        public override Collection<int> GetDirections(MarginEntityInformation marginInformation, MarginMode mode, int marginType)
        {
            Dictionary<int, object> allowedDirections = marginInformation.GetArguments("FabricatrionMargin");          

            Collection<int> directions = new Collection<int>();
            if (allowedDirections != null)
            {
                foreach (object type in allowedDirections.Values)
                {
                    directions.Add(Convert.ToInt32(type));
                }  
            }                               

            return directions;
        }


        /// <summary>
        ///Gets the allowed directions to apply margin by connections. 
        /// </summary>
        /// <param name="connectionMarginInformation">Input connection margin information.</param>
        /// <returns>Returns the list of allowed directions.</returns>
        public override Collection<int> GetDirections(MarginByPartConnectionInformation connectionMarginInformation)
        {
            Dictionary<int, object> allowedDirections = connectionMarginInformation.GetArguments("MarginByConnections");
            Collection<int> directions = new Collection<int>();
            if( allowedDirections!= null)
            {
                foreach (object type in allowedDirections.Values)
                {
                    directions.Add(Convert.ToInt32(type));
                }
            }
            return directions; 
        }


        /// <summary>
        ///Gets the allowed directions for margin by assembly connections.
        /// </summary>
        /// <param name="connectionAssemblyMarginInformation">Input connected assembly information.</param>
        /// <returns>Returns the list of allowed directions.</returns>
        public override Collection<int> GetDirections(MarginByAssemblyConnectionInformation connectionAssemblyMarginInformation)
        {

            Dictionary<int, object> allowedDirections = connectionAssemblyMarginInformation.GetArguments("MarginByConnections");
          
            Collection<int> directions = new Collection<int>();
            if (allowedDirections!= null)
            {
                foreach (object type in allowedDirections.Values)
                {
                    directions.Add(Convert.ToInt32(type));
                }
            }           

            return directions;
        }
   

        /// <summary>
        /// Gets the margin types for fabrication margin.
        /// </summary>
        /// <param name="marginInformation">Input margin information.</param>
        /// <returns>Returns the list of allowed types.</returns>
        public override Collection<int> GetTypes(MarginEntityInformation marginInformation, MarginMode mode)
        {
            if (marginInformation == null)
                throw new CmnNullArgumentException("marginInformation");

            Dictionary<int, object> allowedTypes = null;            

            if(marginInformation.ManufacturingParent is PlatePartBase)
            {
                if (mode == MarginMode.Constant)
                {
                    allowedTypes = marginInformation.GetArguments("PlatePartConstantMargin");
                }
                else if (mode == MarginMode.Oblique)
                {
                    allowedTypes = marginInformation.GetArguments("PlatePartObliqueMargin");
                }                
            }
            else if (marginInformation.ManufacturingParent is ProfilePart)
            {
                allowedTypes = marginInformation.GetArguments("ProfilePartMargin");                
            }

            Collection<int> types = new Collection<int>();
            if (allowedTypes != null)
            {
                foreach (object type in allowedTypes.Values)
                {
                    types.Add(Convert.ToInt32(type));
                }
            }                 

            return types;
        }


        /// <summary>
        /// Gets the margin types for connections by margin.
        /// </summary>
        /// <param name="connectionMarginInformation">Input connection margin information.</param>
        /// <returns>Returns the list of allowed types.</returns>
        public override Collection<int> GetTypes(MarginByPartConnectionInformation connectionByPartEntityInformation)
        {
            Dictionary<int, object> allowedTypes = connectionByPartEntityInformation.GetArguments("MarginByConnection");
            Collection<int> types = new Collection<int>();
            if( allowedTypes!= null)
            {                
                foreach (object type in allowedTypes.Values)
                {
                    types.Add(Convert.ToInt32(type));
                }
            }            
            return types; 
        }


        /// <summary>
        /// Gets the allowed margin types for connections by assembly margin.
        /// </summary>
        /// <param name="connectionbyAssemblyEntityInformation">Input connection assembly information.</param>        
        /// <returns>Returns the list of allowed types.</returns>
        public override Collection<int> GetTypes(MarginByAssemblyConnectionInformation connectionbyAssemblyEntityInformation)
        {
            Dictionary<int, object> allowedTypes = connectionbyAssemblyEntityInformation.GetArguments("MarginByConnection");

            Collection<int> types = new Collection<int>();
            if (allowedTypes != null)
            {
                foreach (object type in allowedTypes.Values)
                {
                    types.Add(Convert.ToInt32(type));
                }
            }           
            return types;       
        }
       

        /// <summary>
        ///Gets the fabrication margin parameters for an individual part. 
        /// </summary>
        /// <param name="marginInformation">Input margin information</param>
        public override MarginParameters GetParameters(MarginEntityInformation marginInformation)
        {
            if (marginInformation == null)
            {
                throw new CmnNullArgumentException("marginInformation");
            }

            PlatePartBase platePart = marginInformation.ManufacturingParent as PlatePartBase;
            ProfilePart profilePart = marginInformation.ManufacturingParent as ProfilePart;
            TopologyPort edgePort = marginInformation.FacePort;

            MarginParameters fabMarginParams = new MarginParameters(1, MarginDirection.NormalToEdge, MarginGeometryOperation.Trim, 0.0, 6.0, 0.0);

            return fabMarginParams;
        }

        /// <summary>
        /// Gets the group margin paramters to apply margin on group of connectd parts. 
        /// </summary>
        /// <param name="connectionMarginInformation">Input connection margin information.</param>
        public override MarginParameters GetParameters(MarginByPartConnectionInformation connectionMarginInformation)
        {

            if (connectionMarginInformation == null)
            {
                throw new CmnNullArgumentException("connectionMarginInformation");
            }

            MarginParameters groupMarginParams = new MarginParameters(1, MarginDirection.NormalToEdge,MarginGeometryOperation.Trim,0.0,15.0,0.0);


            
           /* if( connectionMarginInformation == null)
            {
                throw new CmnNullArgumentException("connectionMarginInformation");
            }

            if( connectionMarginInformation.GroupConnectionType == MarginConnectionType.ReferencePart)
            {
                connectionMarginInformation.GroupType = 1;
                connectionMarginInformation.GroupGeometryOperation = MarginGeometryOperation.Trim;
                connectionMarginInformation.GroupDirection = MarginDirection.NormalToEdge;
                connectionMarginInformation.GroupStretchOffset = 0;
                connectionMarginInformation.GroupMarginStartValue = 15;
            }
            else
            {
                connectionMarginInformation.GroupType = 1;
                connectionMarginInformation.GroupGeometryOperation = MarginGeometryOperation.Trim;
                connectionMarginInformation.GroupDirection = MarginDirection.NormalToEdge;
                connectionMarginInformation.GroupStretchOffset = 0;
                connectionMarginInformation.GroupMarginStartValue = 15;
                connectionMarginInformation.GroupOffsetDirection = base.GetGroupOffsetDirection(connectionMarginInformation.SelectedAssembly, connectionMarginInformation.ConnectedAssembly);                
            }*/


            return groupMarginParams;
        }


        /// <summary>
        /// Gets the group margin paramters to apply margin on connectd assemblies. 
        /// </summary>
        /// <param name="connectionMarginInformation">Input connection margin information.</param>
        public override MarginParameters GetParameters(MarginByAssemblyConnectionInformation connectionAssemblyMarginInformation,out MarginOffsetDirection offsetDirection)
        {

            if (connectionAssemblyMarginInformation == null)
            {
                throw new CmnNullArgumentException("connectionAssemblyMarginInformation");
            }

            MarginParameters groupMarginParams = new MarginParameters(1, MarginDirection.NormalToEdge, MarginGeometryOperation.Trim, 0.0, 12.0, 0.0);
            offsetDirection = MarginOffsetDirection.GlobalX;



            /*if (connectionMarginInformation == null)
            {
                throw new CmnNullArgumentException("connectionMarginInformation");
            }

            if (connectionMarginInformation.GroupConnectionType == MarginConnectionType.ReferencePart)
            {
                connectionMarginInformation.GroupType = 1;
                connectionMarginInformation.GroupGeometryOperation = MarginGeometryOperation.Trim;
                connectionMarginInformation.GroupDirection = MarginDirection.NormalToEdge;
                connectionMarginInformation.GroupStretchOffset = 0;
                connectionMarginInformation.GroupMarginStartValue = 15;
            }
            else
            {
                connectionMarginInformation.GroupType = 1;
                connectionMarginInformation.GroupGeometryOperation = MarginGeometryOperation.Trim;
                connectionMarginInformation.GroupDirection = MarginDirection.NormalToEdge;
                connectionMarginInformation.GroupStretchOffset = 0;
                connectionMarginInformation.GroupMarginStartValue = 15;
                connectionMarginInformation.GroupOffsetDirection = base.GetGroupOffsetDirection(connectionMarginInformation.SelectedAssembly, connectionMarginInformation.ConnectedAssembly);
            }*/

            return groupMarginParams;
        }


        /// <summary>
        /// Gets the individual part margin paramters from group parameters to apply margin.
        /// </summary>
        /// <param name="partMarginInformation">Input part margin information for which margin parameters to define from its group parameters.</param>
        ///  <param name="connectionsByPartMarginInformation">Input connected parts margin entity information.</param>
        /// <param name="groupMarginParameters">Input group margin parameters</param>
        public override MarginParameters GetParameters(MarginEntityInformation partMarginInformation, MarginByPartConnectionInformation connectionsByPartMarginInformation, MarginParameters groupMarginParameters)
        {
            if (partMarginInformation == null)
            {
                throw new CmnNullArgumentException("partMarginInformation");
            }

            if (connectionsByPartMarginInformation == null)
            {
                throw new CmnNullArgumentException("connectionsByPartMarginInformation");
            }

            if (groupMarginParameters == null)
            {
                throw new CmnNullArgumentException("groupMarginParameters");
            }

            MarginParameters marginParams = null;
            if (partMarginInformation.ManufacturingParent is PlatePartBase)
            {
                double stretchOffset = 0.0;
                if (groupMarginParameters.GeometryOperation == MarginGeometryOperation.Stretch)
                {
                    stretchOffset = 2.5;

                }
                marginParams = new MarginParameters(groupMarginParameters.Type, MarginDirection.NormalToEdge, groupMarginParameters.GeometryOperation, stretchOffset, groupMarginParameters.StartValue, groupMarginParameters.EndValue);               
            }
            else if (partMarginInformation.ManufacturingParent is ProfilePart)
            {
                marginParams = new MarginParameters(groupMarginParameters.Type, MarginDirection.NormalToEdge, groupMarginParameters.GeometryOperation, groupMarginParameters.StretchOffset, groupMarginParameters.StartValue, 0.0);              
            }

            return marginParams;
        }


        /// <summary>
        /// Gets the individual part margin paramters from assembly group margin parameters.
        /// </summary>
        /// <param name="partMarginInformation">Input part margin information for which margin parameters to define from its group(assembly) parameters.</param>
        ///  <param name="connectionsByAssemblyMarginInformation">Input connected assembly margin entity information.</param>
        /// <param name="groupMarginParameters">Input group margin parameters</param>
        /// <param name="groupOffsetDirection">Input group offset direction.</param>         

        public override MarginParameters GetParameters(MarginEntityInformation partMarginInformation, MarginByAssemblyConnectionInformation connectionsByAssemblyMarginInformation, MarginParameters groupMarginParameters, MarginOffsetDirection groupOffsetDirection)
        {
            if (partMarginInformation == null)
            {
                throw new CmnNullArgumentException("partMarginInformation");
            }

            if (connectionsByAssemblyMarginInformation == null)
            {
                throw new CmnNullArgumentException("connectionsByAssemblyMarginInformation");
            }

            if (groupMarginParameters == null)
            {
                throw new CmnNullArgumentException("groupMarginParameters");
            }

            MarginParameters marginParams = null;

            if (partMarginInformation.ManufacturingParent is PlatePartBase)
            {
                double stretchOffset = 0.0;
                if (groupMarginParameters.GeometryOperation == MarginGeometryOperation.Stretch)
                {
                    stretchOffset = 2.5;
                }
                marginParams = new MarginParameters(groupMarginParameters.Type, MarginDirection.NormalToEdge, groupMarginParameters.GeometryOperation, stretchOffset, groupMarginParameters.StartValue, 0.0);
             }
            else if (partMarginInformation.ManufacturingParent is ProfilePart)
            {
                marginParams = new MarginParameters(groupMarginParameters.Type, MarginDirection.NormalToEdge, groupMarginParameters.GeometryOperation, groupMarginParameters.StretchOffset, groupMarginParameters.StartValue, 0.0);
            }

            return marginParams;
        }


        /// <summary>
        /// Gets the connected part margin parameters from its parent margin parameters.
        /// </summary>        
        /// <param name="connectedPartMarginInformation">Input connected part's information.</param>       
        /// <param name="parentMarginParameters">Input parent's margin parameters.</param>
        /// <returns>Returns the connected part's margin parameters.</returns>
        public override MarginParameters GetParameters(MarginConnectedPartInformation connectedPartMarginInformation, MarginParameters parentMarginParameters)
        {

            if (connectedPartMarginInformation == null)
            {
                throw new CmnNullArgumentException("connectedPartMarginInformation");
            }
            

            if (parentMarginParameters == null)
            {
                throw new CmnNullArgumentException("parentMarginParameters");
            }

            double stretchOffset = 0.0;
            if (parentMarginParameters.GeometryOperation == MarginGeometryOperation.Stretch)
            {
                stretchOffset = 2.5;

            }
            MarginParameters connectedPartMarginParams = new MarginParameters(parentMarginParameters.Type, MarginDirection.Global, parentMarginParameters.GeometryOperation, stretchOffset, 8.0, 0.0);
            return connectedPartMarginParams;
        }


        /// <summary>
        ///Gets the collection of stiffners that are to update when margin value is changed on a reference part.
        /// </summary>
        /// <param name="partMarginInformation">Input part margin information. which holds the input margin plate.</param>  
        /// <param name="connctedPartsMarginInformation">Input connected part's margin information.</param>         
        /// <returns>Returns the collection of stiffners.</returns>
        public override ReadOnlyCollection<BusinessObject> GetConnectedParts(MarginEntityInformation partMarginInformation, MarginByPartConnectionInformation connctedPartsMarginInformation)
        {

           /* List<string> ObjectNames = new List<string>();

            ObjectNames.Add("<B0>-6LS.2-1");
            ObjectNames.Add("<B0>-0LS.4-1-4LS.1-1");
            ObjectNames.Add("<B0>-0LS.4-1-2LS.1-1");

            Dictionary<string, BusinessObject> objects = null; // ManufacturingATPUtils.GetObjectsFromDatabase(ObjectNames);
            Collection<BusinessObject> connectedObjects = null;
            if( objects != null)
            {
                PlatePartBase platePart = (PlatePartBase)objects["<B0>-6LS.2-1"];
                ProfilePart profile1 = (ProfilePart)objects["<B0>-0LS.4-1-4LS.1-1"];
                ProfilePart profile2 = (ProfilePart)objects["<B0>-0LS.4-1-2LS.1-1"];
                connectedObjects  = new Collection<BusinessObject>();
                connectedObjects.Add(platePart);
                connectedObjects.Add(profile1);
                connectedObjects.Add(profile2);
            }

            if (connectedObjects != null)
                return new ReadOnlyCollection<BusinessObject>(connectedObjects);
            else*/
                return null;
        }


        /// <summary>
        ///Gets the collection of stiffners that are to update when margin value is changed of an referencce assembly's part.
        /// </summary>
        /// <param name="partMarginInformation">Input part margin information. which holds the input margin plate.</param>  
        /// <param name="connectedAssemblyInformation">Input connected assembly information.</param>         
        /// <returns>Returns the collection of stiffners.</returns>
        public override ReadOnlyCollection<BusinessObject> GetConnectedParts(MarginEntityInformation partMarginInformation, MarginByAssemblyConnectionInformation connectedAssemblyInformation)
        {
           /* List<string> ObjectNames = new List<string>();

            ObjectNames.Add("<B0>-6LS.2-1");
            ObjectNames.Add("<B0>-0LS.4-1-4LS.1-1");
            ObjectNames.Add("<B0>-0LS.4-1-2LS.1-1");

            Dictionary<string, BusinessObject> objects = null; // ManufacturingATPUtils.GetObjectsFromDatabase(ObjectNames);
            Collection<BusinessObject> connectedObjects = null;
            if( objects != null)
            {
                PlatePartBase platePart = (PlatePartBase)objects["<B0>-6LS.2-1"];
                ProfilePart profile1 = (ProfilePart)objects["<B0>-0LS.4-1-4LS.1-1"];
                ProfilePart profile2 = (ProfilePart)objects["<B0>-0LS.4-1-2LS.1-1"];
                connectedObjects = new Collection<BusinessObject>();
                connectedObjects.Add(platePart);
                connectedObjects.Add(profile1);
                connectedObjects.Add(profile2);
            }

            return new ReadOnlyCollection<BusinessObject>(connectedObjects);

            if (connectedObjects != null)
            {
               
            }
            else*/

            return null;
        }


        /// <summary>
        /// Gets the behaviour of feature when applied margin on plate part to which this feature belongs.
        /// </summary>
        /// <param name="marginInformation">Input margin information on which feature has to move.</param>
        /// <param name="marginParameters">Input parts margin parameters.</param>
        /// <param name="feature">Input feature of the part which is going under modification as margin is applied on part.</param>
        /// <returns>Returns enumerated feature behaviour.</returns>
        public override MarginFeatureBehavior GetFeatureBehaviour(MarginEntityInformation marginInformation, MarginParameters marginParameters, BusinessObject feature)
        {
            if (feature is Feature == false || feature is Opening == false)
            {
                throw new CmnNullArgumentException("feature");
                //throw new CmnInvalidArgumentException("featureObject");
            }


            MarginFeatureBehavior featureBehaviour = MarginFeatureBehavior.Fix;
            Feature featureObj = feature as Feature;  // successfull means not sketch feature

            //If sketched feature is supplied then code should always return "Fixed Feature" and it should
            //not call get_StructFeatureType as Query Interface for IJStructFeature is not supported

            if (featureObj != null)  // If it is not sketch feature
            {
                if (featureObj.FeatureType == FeatureType.Corner)
                {
                    featureBehaviour = MarginFeatureBehavior.Move;
                }
            }

            return featureBehaviour;
        }
        
        /*




        


        /// <summary>
        /// Gets the individual part margin paramters from group parameters to apply margin.
        /// </summary>
        /// <param name="groupMarginInformation">Input group margin information</param>
        /// <param name="partMarginInformation">Part margin information contains out put margin parameters.</param>       
        public override void GetGroupPartParameters(MarginByConnectionInformation groupMarginInformation, MarginInformation partMarginInformation) //second level
        {

            if( groupMarginInformation == null)
            {
                throw new CmnNullArgumentException("groupMarginInformation");
            }

            try
            {
                if (groupMarginInformation.GroupConnectionType == MarginConnectionType.Assembly)
                {
                    GetAssemblyPartParametersFromGroup(groupMarginInformation, partMarginInformation);
                }
                else
                {
                    GetConnectedPartParametersFromGroup(groupMarginInformation, partMarginInformation);
                }

            }
            catch(Exception e)
            {
                this.WriteToErrorLog(e, "Failed to get base and offset ports " + e.Message, e.Source);
            }
        }


        /// <summary>
        ///Gets the collection of stiffners that are to update when margin value is changed on a part.
        /// </summary>
        /// <param name="groupMarginInformation">Input group margin information.</param>       
        /// <param name="partMarginInformation">Input part margin information. which holds the input margin plate.</param>
        /// <returns>Returns the collection of stiffners.</returns>
        public override ReadOnlyCollection<ProfilePart> GetConnectedStiffners(MarginByConnectionInformation groupMarginInformation,MarginInformation partMarginInformation)
        {
            if (groupMarginInformation == null)
            {
                throw new CmnNullArgumentException("groupMarginInformation");
            }

            if (partMarginInformation == null)
            {
                throw new CmnNullArgumentException("partMarginInformation");
            }

            ReadOnlyCollection<ProfilePart> connectedStiffners = null;
            ReadOnlyCollection<StructPortBase> connections = null;
            Collection<ProfilePart> assemblyConnectedStiffeners = new Collection<ProfilePart>();

            if (groupMarginInformation.GroupConnectionType == MarginConnectionType.ReferencePart)  //ConnectionsByParts Margin
            {
                if (partMarginInformation.ManufacturingParent != null && partMarginInformation.ManufacturingParent is PlatePartBase)
                {
                    connectedStiffners = base.ComputeConnectedStiffeners((PlatePartBase)partMarginInformation.ManufacturingParent);
                }
                else if( partMarginInformation.ManufacturingParent != null && partMarginInformation.ManufacturingParent is ProfilePart)
                {
                    ;
                }                 
            }
            else if (groupMarginInformation.GroupConnectionType == MarginConnectionType.Assembly) //Connections By assembly Margin
            {
                if (groupMarginInformation.SelectedAssembly == null || groupMarginInformation.ConnectedAssembly == null)
                {
                    throw new CmnNullArgumentException("SelectedAssembly or ConnectedAssembly ");
                }

                if (partMarginInformation.ManufacturingParent != null && partMarginInformation.ManufacturingParent is PlatePartBase)
                {
                    ReadOnlyCollection<ProfilePart> connectedStiffnersToAssembly = base.ComputeConnectedStiffeners((PlatePartBase)partMarginInformation.ManufacturingParent);
                    connections = GetConnectionsBetweenAssemblies(groupMarginInformation.SelectedAssembly, groupMarginInformation.ConnectedAssembly);

                    for( int j = 0; j < connections.Count; j++ )
                    {
                        IPort port = connections[j];
                        if(connectedStiffnersToAssembly.Contains( (ProfilePart)port.Connectable))
                        {
                            assemblyConnectedStiffeners.Add((ProfilePart)port.Connectable);
                        }
                    }

                    connectedStiffners = new ReadOnlyCollection<ProfilePart>(assemblyConnectedStiffeners);
                }
            }
            return connectedStiffners;
        }


        /// <summary>
        /// Gets the dependent part (stiffner) margin parameters from parent part for connections margin information.
        /// </summary>
        /// <param name="groupMarginInformation">Input group margin information.</param>
        /// <param name="parentMarginInformation">Input parts parent margin information.</param>
        /// <param name="partMarginInformation">Output dependent parts margin information.</param>        
        public override void GetPartDependentParameters(MarginByConnectionInformation groupMarginInformation,MarginInformation parentMarginInformation, MarginInformation partMarginInformation)
        {
            if (parentMarginInformation == null)
            {
                throw new CmnNullArgumentException("parentMarginInformation");
            }

            if (partMarginInformation == null)
            {
                throw new CmnNullArgumentException("partMarginInformation");

            }             
            if (partMarginInformation.ManufacturingParent == null)
            {
                throw new CmnNullArgumentException("DetailedPart");
            }          

            ProfilePart profile = (ProfilePart)partMarginInformation.ManufacturingParent;

            //Need to review pending

            bool connectionsByAssembly = true;
            if (connectionsByAssembly == true)
            {
                GetProfileParameters(profile, parentMarginInformation);
            }
            else
            {
                GetProfileParametersFromConnectedPlate(profile, parentMarginInformation);
            }

            
        }



        #region Private Helpers

        bool IsPlatePartCanPlate(PlatePartBase platePart)
        {
            bool cylindricalPlate = false; 
            if( platePart.Type == PlateType.TransverseTube || platePart.Type == PlateType.LongitudinalTube || platePart.Type == PlateType.VerticalTube || platePart.Type == PlateType.TubePlate)
            {
                cylindricalPlate = true;
            }

            return cylindricalPlate;
        }

        private MarginInformation GetProfileParameters( ProfilePart profilePart,  MarginInformation plateMarginInfo )
        {
            MarginInformation profilemarginInfo = new MarginInformation(null, "");
            profilemarginInfo.Type = plateMarginInfo.Type;
            profilemarginInfo.Direction = MarginDirection.Global;
            profilemarginInfo.GeometryOperation = plateMarginInfo.GeometryOperation;

            StiffenerPartBase stiffnerProfile = profilePart as StiffenerPartBase;
            Curve3d landingCurve = stiffnerProfile.Axis;

            Position startPos = null;
            Position endPos = null;
            landingCurve.EndPoints(out startPos, out endPos);
            Vector profileDirection = new Vector(endPos.X - startPos.X, endPos.Y - startPos.Y, endPos.Z - startPos.Z);
            profileDirection.Length = 1;

            TopologyPort PlatePort = GetFacePortFromEdgePort(plateMarginInfo.EdgePort);

            Vector portNormal = PlatePort.Normal;
            portNormal.Length = 1;

            double minDistFromStartPos = 0.0;
            double minDisFromEndPos = 0.0;
            Position closestPosFromStart = null;
            Position closestPosFromEnd = null;
            Position closestPos = null;
            Position otherPos = null;
            Point3d startPoint = new Point3d(startPos);
            Point3d endPoint = new Point3d(endPos);

            if (profilemarginInfo.GeometryOperation == MarginGeometryOperation.Stretch)
            {
                PlatePort.DistanceBetween(startPoint, out minDistFromStartPos, out closestPosFromStart);
                PlatePort.DistanceBetween(endPoint, out minDisFromEndPos, out closestPosFromEnd);

                if (minDistFromStartPos < minDisFromEndPos)
                {
                    closestPos = startPos;
                    otherPos = endPos;
                }
                else
                {
                    closestPos = endPos;
                    otherPos = startPos;
                }

                portNormal.Length = plateMarginInfo.StartValue;

                Position newPos = closestPos.Offset(portNormal);

                double oldProfilelength = startPos.DistanceToPoint(endPos);
                double newProfilelength = newPos.DistanceToPoint(otherPos);

                profilemarginInfo.StartValue = Math.Abs(newProfilelength - oldProfilelength);
                profilemarginInfo.EndValue = 0.0;

                TopologyCurve adjacentWireBody = null;
                Position adjacentIntersectPos = null;
                Position adjacentOtherPos = null;

                //Pending
                GetAdjacentPortAndIntersectionPoints(PlatePort, out adjacentWireBody, out adjacentIntersectPos, out adjacentOtherPos);
                Position planePos = null;
                if (plateMarginInfo.StretchOffset > 0)
                {
                    planePos = adjacentWireBody.PointAtDistanceAlong(adjacentIntersectPos, (plateMarginInfo.StretchOffset / 1000));
                    portNormal.Length = 1;

                    IPlane refPlane = new Plane3d(planePos, portNormal);

                    //Pending creation of surfcae from ref plane
                    ISurface surface = null;
                    Collection<BusinessObject> intersectedCurves = null;
                    GeometryIntersectionType intersectionType = GeometryIntersectionType.Unknown;
                    landingCurve.Intersect(surface, out intersectedCurves, out intersectionType);
                    if (intersectedCurves != null)
                    {
                        double distFromStartPt = 0.0;
                        double distFromEndPt = 0.0;
                        Position minPos = null;

                        Point3d closestPt = new Point3d(closestPos);
                        //if( intersectedCurves[0])
                        landingCurve.DistanceBetween(closestPt, out distFromStartPt, out minPos);
                        minPos = null;
                        Point3d planePt = new Point3d(planePos);  //need to get from intersectedCurves[0]

                        landingCurve.DistanceBetween(planePt, out distFromEndPt, out minPos);

                        profilemarginInfo.StretchOffset = distFromStartPt - distFromEndPt;

                        //end if
                        //else
                        Point3d startPt = new Point3d(startPos);
                        landingCurve.DistanceBetween(startPt, out distFromStartPt, out minPos);
                        profilemarginInfo.StretchOffset = distFromStartPt;

                        profilemarginInfo.StretchOffset = profilemarginInfo.StretchOffset * 1000;  // return in mm


                    }
                }
                else
                {
                    profilemarginInfo.StretchOffset = 0.0;
                }
            }
            else
            {
                profilemarginInfo.StartValue = 0.0;
               if( Math.Abs(portNormal.Dot(profileDirection))> 0 )
               {
                   profilemarginInfo.StartValue = plateMarginInfo.StartValue / Math.Abs(portNormal.Dot(profileDirection));
               }
               profilemarginInfo.EndValue = 0.0;
               profilemarginInfo.StretchOffset = 0.0;              

            }
            return null;
        }


        private MarginInformation GetProfileParametersFromConnectedPlate(ProfilePart profilePart, MarginInformation plateMarginInfo)
        {

            MarginInformation profilemarginInfo = new MarginInformation(null, "");

            profilemarginInfo.Type = plateMarginInfo.Type;
            profilemarginInfo.Direction = MarginDirection.Global;
            profilemarginInfo.GeometryOperation = plateMarginInfo.GeometryOperation;

            StiffenerPartBase stiffnerProfile = profilePart as StiffenerPartBase;
            Curve3d landingCurve = stiffnerProfile.Axis;

            Position profileCurveStartPos = null;
            Position profileCurveEndPos = null;
            landingCurve.EndPoints(out profileCurveStartPos, out profileCurveEndPos);

            Vector profileDirection = new Vector(profileCurveEndPos.X - profileCurveStartPos.X, profileCurveEndPos.Y - profileCurveStartPos.Y, profileCurveEndPos.Z - profileCurveStartPos.Z);
            profileDirection.Length = 1;

            TopologyPort PlatePort = GetFacePortFromEdgePort(plateMarginInfo.EdgePort);
            Vector portNormal = PlatePort.Normal;
            portNormal.Length = 1;

            double minDistFromStartPos = 0.0, minDisFromEndPos = 0.0;
            Position closestPosFromStart = null, closestPosFromEnd = null;
            Position closestPos = null,otherPos = null;

            Point3d startPoint = new Point3d(profileCurveStartPos);
            Point3d endPoint = new Point3d(profileCurveEndPos);

            if (profilemarginInfo.GeometryOperation == MarginGeometryOperation.Stretch)
            {
                PlatePort.DistanceBetween(startPoint, out minDistFromStartPos, out closestPosFromStart);
                PlatePort.DistanceBetween(endPoint, out minDisFromEndPos, out closestPosFromEnd);
                if (minDistFromStartPos < minDisFromEndPos)
                {
                    closestPos = profileCurveStartPos;
                    otherPos = profileCurveEndPos;
                }
                else
                {
                    closestPos = profileCurveEndPos;
                    otherPos = profileCurveStartPos;
                }

                portNormal.Length = plateMarginInfo.StartValue;
                Position newPos = closestPos.Offset(portNormal);

                double oldProfileLength = profileCurveStartPos.DistanceToPoint(profileCurveEndPos);
                double newprofileLength = newPos.DistanceToPoint(otherPos);

                profilemarginInfo.StartValue = Math.Abs(newprofileLength - oldProfileLength);
                profilemarginInfo.EndValue = 0.0;
                profilemarginInfo.StretchOffset = 0.0;

            }
            else
            {
                profilemarginInfo.StartValue = plateMarginInfo.StartValue / Math.Abs(portNormal.Dot(profileDirection));
                profilemarginInfo.EndValue = 0.0;
                profilemarginInfo.StretchOffset = 0.0;
            }

            return null;
        }


        private void GetAdjacentPortAndIntersectionPoints( IPort facePort,out TopologyCurve adjacentWireBody, out Position intersectionPos1,out Position intersectionpos2)
        {
            adjacentWireBody = null;
            intersectionPos1 = null;
            intersectionpos2 = null;

             IConnectable connectable = facePort.Connectable;
             ReadOnlyCollection<IPort> connectedPorts =  connectable.GetConnectedPorts(PortType.Face);

             TopologyCurve marginEdgeBody = null;
             TopologyCurve otherEdgeBody = null;

            //Pending
             marginEdgeBody = GetBaseEdgeFromFacePort(facePort);

            for( int  i = 0; i < connectedPorts.Count; i++)
            {
                IPort otherPort = connectedPorts[i];

                TopologyPort structPort = otherPort as TopologyPort;
                if( structPort.ContextId == ContextTypes.Lateral )
                {
                    otherEdgeBody = GetBaseEdgeFromFacePort(otherPort);
                }

                if( otherPort != facePort)
                {
                    Collection<Position>intersecPosColl = null;
                    Collection<Position>overlapPosColl = null;
                    Position intersectPos = null;
                    GeometryIntersectionType intersectionType = GeometryIntersectionType.Unknown;

                    ICurve marginCurve = marginEdgeBody as ICurve;
                    ICurve otherCurve = (ICurve)otherEdgeBody;
                    marginCurve.Intersect(otherCurve, out intersecPosColl, out overlapPosColl, out intersectionType);

                    if( intersecPosColl != null && intersecPosColl.Count > 0 )
                    {
                        intersectPos = intersecPosColl[0];
                    }

                    Position projectedPos = otherCurve.ProjectPoint(intersectPos);

                    Position otherWBStartPos = null, otherWBEndPos = null;
                    otherCurve.EndPoints(out otherWBStartPos, out otherWBEndPos);

                    double dist1 = otherWBStartPos.DistanceToPoint(projectedPos);
                    double dist2 = otherWBEndPos.DistanceToPoint(projectedPos);
                    if( dist1 < dist2)
                    {
                        intersectionpos2 = otherWBEndPos;
                        intersectionPos1 = otherWBStartPos;
                    }
                    else
                    {
                        intersectionpos2 = otherWBStartPos;
                        intersectionPos1 = otherWBEndPos;
                    }

                    adjacentWireBody = otherEdgeBody;
                    break;
                }                
            }
        }


        private void GetAssemblyPartParametersFromGroup(MarginByConnectionInformation groupMarginInformation, MarginInformation partMarginInformation)
        {

            if(partMarginInformation.ManufacturingParent is PlatePartBase )
            {
                PlatePart part = partMarginInformation.ManufacturingParent as PlatePart;
                if( part.IsRoot == true)
                {
                    PlateSystemBase system = part.RootPlateSystem;

                    //Codereview: need to use functions on platePart for tripping or plane
                    if(base.IsPlaneBracket(system) || base.IsTrippingBracket(system) )
                    {
                        partMarginInformation.Mode = MarginMode.Constant;
                        partMarginInformation.StartValue = 0;
                        partMarginInformation.GeometryOperation = MarginGeometryOperation.Trim;
                        partMarginInformation.Direction = MarginDirection.NormalToEdge;
                        partMarginInformation.StretchOffset = 0;
                        return;
                    }
                }

                CustomPlatePart smartPlate = partMarginInformation.ManufacturingParent as CustomPlatePart;
                CollarPart collarPlate = partMarginInformation.ManufacturingParent as CollarPart;
                if( smartPlate != null || collarPlate != null)
                {
                    partMarginInformation.Mode = MarginMode.Constant;
                    partMarginInformation.StartValue = 0;
                    partMarginInformation.GeometryOperation = MarginGeometryOperation.Trim;
                    partMarginInformation.Direction = MarginDirection.NormalToEdge;
                    partMarginInformation.StretchOffset = 0;
                    return;
                }
            }

            partMarginInformation.Mode = MarginMode.Assembly;
            if (partMarginInformation.ManufacturingParent is PlatePartBase)
            {
                GetPlateMarginParameters(partMarginInformation, groupMarginInformation);                
            }
            else if (partMarginInformation.ManufacturingParent is ProfilePart)
            {  
                PlatePartBase stiffenedPlate = null;
                TopologyPort stiffenedEdge  =null;

                base.GetStiffenedPlatePart((ProfilePart)partMarginInformation.ManufacturingParent, groupMarginInformation.SelectedAssembly, groupMarginInformation.ConnectedAssembly, out stiffenedPlate, out stiffenedEdge);
                if( stiffenedPlate == null)
                {
                    partMarginInformation.Type = 1;
                    partMarginInformation.GeometryOperation = MarginGeometryOperation.Trim;
                    partMarginInformation.StartValue = 10.0;
                    partMarginInformation.Direction = MarginDirection.Global;
                    return;
                }

                MarginInformation stiffenedPlateMarginInfo = new MarginInformation(null,"");
                stiffenedPlateMarginInfo.ManufacturingParent = (IManufacturable) stiffenedPlate;
                stiffenedPlateMarginInfo.EdgePort = stiffenedEdge;

                GetPlateMarginParameters(stiffenedPlateMarginInfo, groupMarginInformation);

                GetProfileParameters((ProfilePart)partMarginInformation.ManufacturingParent, stiffenedPlateMarginInfo);
            }

        }


        private void GetPlateMarginParameters(MarginInformation PlatePartMarginInformation,MarginByConnectionInformation groupMarginInformation)
        {
            PlatePartMarginInformation.Type = groupMarginInformation.GroupType;
            PlatePartMarginInformation.Direction = MarginDirection.NormalToEdge;
            if (groupMarginInformation.GroupConnectionType == MarginConnectionType.Assembly)
                PlatePartMarginInformation.StartValue = GetMarginValueBasedOnOffsetDirection(PlatePartMarginInformation.EdgePort,groupMarginInformation.GroupOffsetDirection,groupMarginInformation.GroupMarginStartValue);
            else
            {
                PlatePartMarginInformation.StartValue = groupMarginInformation.GroupMarginStartValue;
            }
            PlatePartMarginInformation.EndValue = 0.0;
            PlatePartMarginInformation.GeometryOperation = groupMarginInformation.GroupGeometryOperation;
            PlatePartMarginInformation.StretchOffset = 0.0;
            if( PlatePartMarginInformation.GeometryOperation == MarginGeometryOperation.Stretch )
            {
                PlatePartMarginInformation.StretchOffset = ComputeStrechOffset((PlatePartBase)PlatePartMarginInformation.ManufacturingParent, PlatePartMarginInformation.EdgePort, groupMarginInformation.GroupStretchOffset);
            }
        }


        private void GetConnectedPartParametersFromGroup(MarginByConnectionInformation groupMarginInformation, MarginInformation partMarginInformation)
        {
            if (partMarginInformation.ManufacturingParent is PlatePartBase)
            {
                PlatePart part = partMarginInformation.ManufacturingParent as PlatePart;
                if (part.IsRoot == true)
                {
                    PlateSystemBase system = part.RootPlateSystem;

                    //Codereview: need to use functions on platePart for tripping or plane
                    if (base.IsPlaneBracket(system) || base.IsTrippingBracket(system))
                    {
                        partMarginInformation.Mode = MarginMode.Constant;
                        partMarginInformation.Type = 1;
                        partMarginInformation.StartValue = 10;
                        partMarginInformation.GeometryOperation = MarginGeometryOperation.Trim;
                        partMarginInformation.Direction = MarginDirection.NormalToEdge;
                        partMarginInformation.StretchOffset = 0;
                        return;
                    }
                }

                CustomPlatePart smartPlate = partMarginInformation.ManufacturingParent as CustomPlatePart;
                CollarPart collarPlate = partMarginInformation.ManufacturingParent as CollarPart;
                if (smartPlate != null || collarPlate != null)
                {
                    partMarginInformation.Mode = MarginMode.Constant;
                    partMarginInformation.Type = 1;
                    partMarginInformation.StartValue = 10;
                    partMarginInformation.GeometryOperation = MarginGeometryOperation.Trim;
                    partMarginInformation.Direction = MarginDirection.NormalToEdge;
                    partMarginInformation.StretchOffset = 0;
                    return;
                }
            }

            partMarginInformation.Mode = MarginMode.Constant;
            if (partMarginInformation.ManufacturingParent is PlatePartBase)
            {
                GetPlateMarginParameters(partMarginInformation, groupMarginInformation);
            }
            else if (partMarginInformation.ManufacturingParent is ProfilePart)
            {
                PlatePartBase stiffenedPlate = null;
                TopologyPort stiffenedEdge = null;

                //base.GetStiffenedPlatePartFromConnectedPlates((ProfilePart)partMarginInformation.ManufacturingParent, groupMarginInformation.SelectedAssembly, groupMarginInformation.ConnectedAssembly, out stiffenedPlate, out stiffenedEdge);
                if (stiffenedPlate == null || stiffenedEdge == null)
                {
                    partMarginInformation.Type = groupMarginInformation.GroupType;
                    partMarginInformation.GeometryOperation = groupMarginInformation.GroupGeometryOperation;
                    partMarginInformation.StartValue = groupMarginInformation.GroupMarginStartValue;
                    partMarginInformation.Direction = MarginDirection.NormalToEdge;
                    partMarginInformation.StretchOffset = groupMarginInformation.GroupStretchOffset;                   
                }
                else
                {
                    MarginInformation stiffenedPlateMarginInfo = new MarginInformation(null, "");
                    stiffenedPlateMarginInfo.ManufacturingParent = (IManufacturable)stiffenedPlate;
                    stiffenedPlateMarginInfo.EdgePort = stiffenedEdge;
                    GetPlateMarginParameters(stiffenedPlateMarginInfo, groupMarginInformation);
                    GetProfileParametersFromConnectedPlate((ProfilePart)partMarginInformation.ManufacturingParent, stiffenedPlateMarginInfo);
                }
            }
        }*/

        //#endregion Private Helpers



        /*

        /// <summary>
        /// Gets the allowed directions for fabrication margin.
        /// </summary>
        /// <param name="activeMarginInfo">Transient class which holds the input arguments necessary to this function execution.</param>
        /// <param name="mode">Indicates the margin mode that is being choosen. </param>
        /// <param name="marginType">Indicates the type of margin being applied.</param>
        /// <returns>Returns collection of allowed directions.</returns>
        public override ReadOnlyCollection<int> GetDirections(MarginInformation activeMarginInfo)
        {

            Dictionary<int, object> allowedDirections = activeMarginInfo.GetArguments("Margin");

            ArrayList allowableList = new ArrayList(allowedDirections.Values);

            Collection<int> directions = new Collection<int>();

            for (int idx = 0; idx < allowableList.Count; idx++)
            {
                directions.Add(Convert.ToInt32(allowableList[idx]));
            }            

            return new ReadOnlyCollection<int> (directions);
        }


        /// <summary>
        /// Gets the allowed directions for connection by parts margin.
        /// </summary>
        /// <param name="referenceMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns collection of allowed directions.</returns>
        public override ReadOnlyCollection<int> GetDirections(ReferencePartConnectionsMarginInformation referenceMarginInfo)
        {
            return null;

        }


        /// <summary>
        /// Gets the allowed directions for connection by assembly margin.
        /// </summary>
        /// <param name="assemblyMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns collectiion of allowed directions.</returns>
        public override ReadOnlyCollection<int> GetDirections(AssemblyConnectionMarginInformation assemblyMarginInfo)
        {
            return null;

        }


        /// <summary>
        ///  Gets the margin types for fabrication margin.
        /// </summary>
        /// <param name="activeMarginInfo">Transient class which holds the input arguments necessary to this function execution</param>
        /// <param name="mode">Indicates the margin mode that is being choosen.</param>
        /// <returns>Returns collection of allowed types.</returns>

        public override ReadOnlyCollection<int> GetTypes(ActiveMarginInformation activeMarginInfo, MarginMode mode)
        {
            return null;

        }


        /// <summary>
        /// Gets the margin types for connection by parts margin.
        /// </summary>
        /// <param name="referenceMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param> 
        /// <returns>Returns collection of allowed types.</returns>
        public override ReadOnlyCollection<int> GetTypes(ReferencePartConnectionsMarginInformation referenceMarginInfo)
        {
            return null;

        }



        /// <summary>
        /// Gets the margin types for connections by assembly margin.
        /// </summary>
        /// <param name="assemblyMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param> 
        /// <returns>Returns collection of allowed types.</returns>
        public override ReadOnlyCollection<int> GetTypes(AssemblyConnectionMarginInformation assemblyMarginInfo)
        {
            return null;

        }


        /// <summary>
        ///  Gets the fabrication margin parameters for a part.
        /// </summary>
        /// <param name="activeMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns margin parameters for a part.</returns>
        public override MarginParameters GetParameters(ActiveMarginInformation activeMarginInfo)
        {
            return null;

        }

        /// <summary>
        ///  Gets the part margin parameters from group margin parameters for connection by parts margin.
        /// </summary>
        /// <param name="referenceMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns margin parameters for a part that is connected to a reference part.</returns>
        public override MarginParameters GetParameters(ReferencePartConnectionsMarginInformation referenceMarginInfo)
        {
            return null;

        }



        /// <summary>
        ///  Gets the part margin parameters from assembly margin parameters for connection by assembly margin.
        /// </summary>
        /// <param name="assemblyMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns margin parameters for a part that is child of an assembly.</returns>
        public override MarginParameters GetParameters(AssemblyConnectionMarginInformation assemblyMarginInfo)
        {
            return null;

        }

        
        /// <summary>
        /// Gets the behaviour of feature when applied margin on plate part to which this feature belongs.
        /// </summary>
        /// <param name="activeMarginInfo"></param>
        /// <returns>Transient class which holds the input and out put arguments necessary to this function execution.</returns>
        public override MarginFeatureBehavior GetFeatureBehaviour(ActiveMarginInformation activeMarginInfo)
        {
            return MarginFeatureBehavior.Move;

        }


        /// <summary>
        ///  Gets the group margin parametrs for all connected parts for connections by parts margin.
        /// </summary>
        /// <param name="referenceMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns margin parameters for a group of parts.</returns>
        public override MarginParameters GetGroupParameters(ReferencePartConnectionsMarginInformation referenceMarginInfo)
        {
            return null;

        }


        /// <summary>
        /// Gets the group margin parameters for an assembly for connections by assembly margin.
        /// </summary>
        /// <param name="assemblyMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns margin parameters for all childs of an assembly .</returns>
        public override MarginParameters GetGroupParameters(AssemblyConnectionMarginInformation assemblyMarginInfo)
        {
            return null;

        }


        /// <summary>
        ///  Gets the collectiion stiffners that are to update when margin is changed on a part.
        /// </summary>
        /// <param name="referenceMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns collection of connected stiffners.</returns>
        public override ReadOnlyCollection<StiffenerPartBase> GetConnectedStiffeneresToUpdate(ReferencePartConnectionsMarginInformation referenceMarginInfo)
        {
            return null;
            
        }


        /// <summary>
        ///Gets the collection of stiffners to be updated for connections by assemblies when one of the 
        /// assembly's part margin is changed.
        /// </summary>
        /// <param name="assemblyMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns>Returns collection of connected stiffners.</returns>
        public override ReadOnlyCollection<StiffenerPartBase> GetConnectedStiffeneresToUpdate(AssemblyConnectionMarginInformation assemblyMarginInfo)
        {
            return null;

        }


        /// <summary>
        ///  Gets the dependent stiffner margin parameters from parent part for connections by parts margin.
        /// </summary>
        /// <param name="referenceMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns> Returns dependent margin parameters.</returns>
        public override MarginParameters GetDependentStiffenerParameters(ReferencePartConnectionsMarginInformation referenceMarginInfo)
        {
            return null;

        }


        /// <summary>
        ///  Gets the dependent stiffner margin parameters from parent part for connections by assembly margin.
        /// </summary>
        /// <param name="assemblyMarginInfo">Transient class which holds the input and out put arguments necessary to this function execution.</param>
        /// <returns> Returns dependent margin parameters.</returns>
        public override MarginParameters GetDependentStiffenerParameters(AssemblyConnectionMarginInformation assemblyMarginInfo)
        {
            return null;

            
        }
          
         */

    }
}
