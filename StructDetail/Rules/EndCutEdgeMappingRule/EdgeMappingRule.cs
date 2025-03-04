//-----------------------------------------------------------------------------
//      Copyright (C) 2010-16 Intergraph Corporation.  All rights reserved.
//
//      Component:  EdgeMappingRule class is designed to be the user customizable class
//                  which will return the appropriate cross section mapping.
//
//      Author:  3XCalibur
//
//      History:
//      November 03, 2010   WR              Created
//      December 27th, change made to exit the mapping rule in cases of FreeEndCuts, end-to-end connections(include splits, miters) as part of 207893.
//      October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) 'GetSectionEdges' method is called with new parameter 'pointMap' 
//      November 29,2014    Alliagtors      DI-259276 Replace Interop call to mapping utility with .NET API equivalent.
//		November 02,2015	knukala			DI-CP-275245,275246,275247,275248 Removed Introp calls
//      January  25,2016    pkakula         TR-286738: Updated GetEdgeMapping()
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;
using System.Configuration;
using System.Collections;
using System;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Structure;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Content.Structure;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// End cut edge mapping rule gets the given bounded/bounding information and returns the mapping information.
    /// </summary>
    [WrapperProgID("MarineRuleWrappers.EndCutMappingWrapper")]
    public class EdgeMappingRule : EndCutMappingRuleBase, ICustomEdgeMappingRule
    {
        #region ICustomEdgeMappingRule Members

        /// <summary>
        /// Gets the edge mapping.
        /// </summary>
        /// <param name="boundingPort">The bounding port.</param>
        /// <param name="boundedPort">The bounded port.</param>
        /// <param name="sectionAlias">The section alias.</param>
        /// <param name="penetratesWeb">if set to <c>true</c> [penetrates web].</param>
        /// <returns>Edge cut mapping information.</returns>
        /// <exception cref="ArgumentNullException">The passed port is null.</exception>
        public override Dictionary<int, int> GetEdgeMapping(IPort boundingPort, IPort boundedPort, out int sectionAlias, out bool penetratesWeb)
        {
            if (boundingPort == null)
            {
                throw new ArgumentNullException("boundingPort");
            }

            if (boundedPort == null)
            {
                throw new ArgumentNullException("boundedPort");
            }

            // If not a member, then return an empty map
            Dictionary<int, int> edges = new Dictionary<int, int>();
            //check to see if its a free end cut
            if (boundedPort == boundingPort)
            {
                sectionAlias = (int)SectionAliasType.UnknownAlias;
                penetratesWeb = false;
                return edges;

            }
            Ingr.SP3D.Structure.Middle.ProfilePart boundingMember = boundingPort.Connectable as Ingr.SP3D.Structure.Middle.ProfilePart;
            if (boundingMember == null)
            {
                sectionAlias = (int)SectionAliasType.UnknownAlias;
                penetratesWeb = false;
                return edges;
            }

            Ingr.SP3D.Structure.Middle.ProfilePart boundedmember = boundedPort.Connectable as Ingr.SP3D.Structure.Middle.ProfilePart;
            if (boundedmember == null)
            {
                sectionAlias = (int)SectionAliasType.UnknownAlias;
                penetratesWeb = false;
                return edges;
            }

            StiffenerPartBase boundingStiffenr = boundingPort.Connectable as StiffenerPartBase;
            StiffenerPartBase boundedStiffenr = boundedPort.Connectable as StiffenerPartBase;

            // For Stiffener - Stiffener case, check we need to go for Mapping for just Exit
            if (boundedStiffenr != null && boundingStiffenr != null)
            {
                // Get the AssemblyConnection between Stiffeners
                BusinessObject assemblyConnection = GetAssemblyConnection(boundingPort, boundedPort);
                bool bDoMapping = false;

                if (assemblyConnection == null)
                {
                    // NO Assembly Conenction
                    // Get the Primary Assembly Connection to find if the bounding Stiffener is part of Auxiliary Port connection
                    assemblyConnection = GetPrimaryAssemblyConnection(boundingPort, boundedPort);

                    // The bounding Stiffener is part of Auxiliary connection, do the Mapping or else Exit
                    if (assemblyConnection != null)
                        bDoMapping = true;
                }
                else
                {
                    // If the connection has Auxiliary ports, do the Mapping, or else Exit
                    if (HasAuxiliaryPorts(assemblyConnection) == true)
                        bDoMapping = true;

                }

                if (bDoMapping == false)
                {
                    sectionAlias = (int)SectionAliasType.UnknownAlias;
                    penetratesWeb = false;
                    return edges;
                }
            }

            S3DEdgeMap edgeMap = new S3DEdgeMap(boundingPort, boundedPort);
            penetratesWeb = edgeMap.IsWebPenetrating();

            //check made to prvent the call to Getsectionedges method whenever bounding is a designed member (Since currently builtup is not handles as member).

            MemberPart boundingmemberpart = boundingPort.Connectable as Ingr.SP3D.Structure.Middle.MemberPart;

            if (boundingmemberpart != null && (boundingmemberpart.DesignedMember))
            {
                sectionAlias = (int)SectionAliasType.UnknownAlias;
                return edges;
            }

            Dictionary<SectionFaceType, SectionFaceType> mappedEdges = new Dictionary<SectionFaceType, SectionFaceType>();
            string sectionType = boundingMember.SectionType;

            //change made to exit edgemapping rule when it is end-to-end connection<start>

            Boolean bBoundedBase = false;
            Boolean bBoundedOffset = false;
            Boolean bBoundingBase = false;
            Boolean bBoundingOffset = false;


            if (boundedPort is MemberPartAxisPort)
            {
                MemberPartAxisPort oBoundedMemberPort = boundedPort as MemberPartAxisPort;

                if (oBoundedMemberPort != null)
                {
                    if (oBoundedMemberPort.AxisPortType == MemberAxisPortType.Start)
                    {
                        bBoundedBase = true;
                    }
                    else if (oBoundedMemberPort.AxisPortType == MemberAxisPortType.End)
                    {
                        bBoundedOffset = true;
                    }
                }
            }
            else
            {
                TopologyPort oBoundedTopoPort = boundedPort as TopologyPort;

                if (oBoundedTopoPort != null)
                {
                    if (oBoundedTopoPort.ContextId == ContextTypes.Base)
                    {
                        bBoundedBase = true;
                    }
                    else if (oBoundedTopoPort.ContextId == ContextTypes.Offset)
                    {
                        bBoundedOffset = true;
                    }
                }
            }
            MemberPartAxisPort oBoundingMemberPort = boundingPort as MemberPartAxisPort;

            if (oBoundingMemberPort != null)
            {
                if (oBoundingMemberPort.AxisPortType == MemberAxisPortType.Start)
                {
                    bBoundingBase = true;
                }
                else if (oBoundingMemberPort.AxisPortType == MemberAxisPortType.End)
                {
                    bBoundingOffset = true;
                }
            }

            if ((bBoundedBase == true) || (bBoundedOffset == true))
            {
                if ((bBoundingBase == true) || (bBoundingOffset == true))
                {
                    sectionAlias = (int)SectionAliasType.UnknownAlias;
                    penetratesWeb = false;
                    return edges;
                }
            }
            //change made to exit edgemapping rule when it is end-to-end connection.<end>

            int sectionQuadrant = edgeMap.GetOrientationQuadrant();
            bool flipLeftAndRight = edgeMap.IsPenetrationWithSketchingPlaneNormal();
            Dictionary<int, int> pointMap = new Dictionary<int, int>();

            string xmlMapFileLocation = ConfigurationManager.AppSettings["XMLMapFileLocation"];
            if (string.IsNullOrEmpty(xmlMapFileLocation))
            {
                mappedEdges = edgeMap.GetSectionEdges(sectionType, flipLeftAndRight, sectionQuadrant, pointMap);

                bool isNullOrEmpty = true;
                if (mappedEdges != null)
                {
                    if (mappedEdges.Count > 0)
                    {
                        isNullOrEmpty = false;
                    }
                }

                if (isNullOrEmpty)
                {
                    throw new NullReferenceException("Unable to retrieve the end cut edge mapping info.");
                }
            }
            else
            {
                //TODO:  Add code here to get the XML maps.
                throw new NotImplementedException();
            }

            // get the section alias based on the edge mapping obtained above
            sectionAlias = (int)SectionAlias.GetSectionAlias(mappedEdges);

            foreach (var edge in mappedEdges)
            {
                edges.Add((int)edge.Key, (int)edge.Value);
            }

            // REMAP: <start>---from here consider for remapping based on bounding to Bounded Orientation

            bool bIsReMapped;
            int newQuadrant;

            edgeMap.ReMapIfNeeded(sectionQuadrant, sectionAlias, edges, out newQuadrant, out bIsReMapped);
            if (bIsReMapped)
            {
                mappedEdges.Clear();
                edges.Clear();
                pointMap.Clear();
                mappedEdges = edgeMap.GetSectionEdges(sectionType, flipLeftAndRight, newQuadrant, pointMap);

                bool isNullOrEmpty = true;
                if (mappedEdges != null)
                {
                    if (mappedEdges.Count > 0)
                    {
                        isNullOrEmpty = false;
                    }
                }

                if (isNullOrEmpty)
                {
                    throw new NullReferenceException("Unable to retrieve the end cut edge mapping info during remap.");
                }

                // get the section alias based on the edge mapping obtained above
                sectionAlias = (int)SectionAlias.GetSectionAlias(mappedEdges);
                penetratesWeb = edgeMap.IsWebPenetrating();

                foreach (var edge in mappedEdges)
                {
                    edges.Add((int)edge.Key, (int)edge.Value);
                }
            }

            // REMAP: <end>----from here consider for remapping based on bounding to Bounded Orientation

            foreach (var pointID in pointMap)
            {
                edges.Add(pointID.Key, pointID.Value);
            }

            return edges;
        }

        // **********************************************************************************************************
        // * THE FOLLOWING METHODS ARE TO BE IMPLEMENTED BY ALLIGATORS
        // **********************************************************************************************************
        /// <summary>
        /// Gets edge mapping for an end cut with multiple bounding objects.
        /// </summary>
        /// <param name="feature">The end cut.</param>
        /// <param name="sketchingPlane">The end cut sketching plane.</param>
        /// <returns> A collection indicating the relationship between symbol edge IDs and bounding ports.
        /// The keys are special edge IDs used for multiple-boundary end cut symbols.
        /// From top/left to bottom/right the values are 5001, 5002, 5003...5000+n.
        /// The values are the ports that should be used when re-symbolizing the edges.</returns>
        /// <remarks>Restrictions:
        /// 1) The boundaries cannot fully overlap
        /// 2) The bounding parts must fully overlap the end of the bounding member.
        /// If all restrictions are not met, an empty collection is returned.</remarks>
        public override Dictionary<int, TopologyPort> GetMultipleBoundaryEdgeMapping(Feature feature, IPlane sketchingPlane)
        {
            EndcutMappingPlaneOption mappingPlaneOption;

            if (feature.FeatureType == FeatureType.WebCut)
                mappingPlaneOption = EndcutMappingPlaneOption.WebMidThickness;
            else if (feature.FeatureType == FeatureType.FlangeCut)
            {
                //Check if top flange
                bool topFlange = feature.IsTopFlangeCut;

                if (topFlange)
                {
                    mappingPlaneOption = EndcutMappingPlaneOption.TopFlangeMidThickness;
                }
                else
                {
                    mappingPlaneOption = EndcutMappingPlaneOption.BottomFlangeMidThickness;
                }
            }
            else
            {
                throw new Exception("Feature is neither a webcut nor flangecut ");
            }

            Dictionary<int, TopologyPort> sequencedMappedPorts = new Dictionary<int, TopologyPort>();
            AssemblyConnection assemblyConnection = GetParentAssemblyConnection(feature);

            System.Collections.ObjectModel.Collection<IPort> boundedPortCollection = new System.Collections.ObjectModel.Collection<IPort>();
            System.Collections.ObjectModel.Collection<IPort> boundingPortCollection = new System.Collections.ObjectModel.Collection<IPort>();


            if (assemblyConnection != null)
            {
                boundedPortCollection = assemblyConnection.BoundedPorts;
                boundingPortCollection = assemblyConnection.BoundingPorts;
            }
            else
            {
                throw new Exception("AssemblyConnection");
            }

            if (boundedPortCollection.Count != 1)
            {
                throw new Exception("bounded ports count is not eqaul to one");
            }

            IPort boundedPort = boundedPortCollection[0];

            try
            {

                //Get the Mapped Data
                OrderedEndCutPortsInformation OrderedEndCutPortsData = GetMultipleBoundaryEdgeMapping(boundingPortCollection, boundedPortCollection[0], mappingPlaneOption);

                IDictionary<int, IPort> orderedEndcutPorts = OrderedEndCutPortsData.OrderedEndCutPorts;

                if (orderedEndcutPorts != null)
                {
                    foreach (var map in orderedEndcutPorts)
                    {
                        sequencedMappedPorts.Add(map.Key, (TopologyPort)map.Value);
                    }
                }
            }
            catch
            {
                throw new Exception("OrderedEndCutPortsData");
            }
            return sequencedMappedPorts;
        }

        /// <summary>
        /// Returns sequenced edge mapping for a series of adjacent objects to be used as boundaries for a single profile. 
        /// </summary>
        /// <param name="boundingPorts">The full list of ports selected by the user to bound the member.  
        /// Only one port per object is required, and the port does not need to be intersected by the bounded object.
        /// The rule will determine the actual ports needed to bound the member and return these. </param>
        /// <param name="boundedPort">The bounded Port.</param>
        /// <param name="mappingPlaneOption">Option indicating which sketching plane shall be used for the mapping process.</param>
        /// <returns>Edge cut mapping information.</returns>
        /// <remarks>Restrictions:
        /// 1) The boundaries cannot fully overlap
        /// 2) The bounding parts must fully overlap the end of the bounding member.
        /// If all restrictions are not met, an empty collection is returned.</remarks>
        public override OrderedEndCutPortsInformation GetMultipleBoundaryEdgeMapping(IEnumerable<IPort> boundingPorts, IPort boundedPort, EndcutMappingPlaneOption mappingPlaneOption)
        {
            OrderedEndCutPortsInformation OrderedEndCutPortsData = null;

            if (boundedPort == null)
            {
                throw new ArgumentNullException("boundedPort");
            }

            if (boundingPorts == null)
            {
                throw new ArgumentNullException("boundingPorts");
            }

            //eEndCutTypes forEndcutType;
            EndCutType eEndcutType;
            if (mappingPlaneOption == EndcutMappingPlaneOption.WebMidThickness) eEndcutType = EndCutType.WebCut;
            else if (mappingPlaneOption == EndcutMappingPlaneOption.TopFlangeMidThickness) eEndcutType = EndCutType.FlangeCutTop;
            else if (mappingPlaneOption == EndcutMappingPlaneOption.BottomFlangeMidThickness) eEndcutType = EndCutType.FlangeCutBottom;
            else
            {
                throw new ArgumentOutOfRangeException("mappingPlaneOption");
            }

            IDictionary<int, IPort> orderedEndCutPorts = null;

            OrderedEndCutPortsInformation orderedEndcutPortsInformation = Feature.GetOrderedEndCutPorts(boundedPort, boundingPorts, eEndcutType, 1e-005);
            orderedEndCutPorts = orderedEndcutPortsInformation.OrderedEndCutPorts;

            // create the MultipleBoundaryEdgeMapping to hold the data
            if (orderedEndCutPorts != null)
            {
                Dictionary<int, IPort> sequencedMappedPorts = new Dictionary<int, IPort>();

                int edgeID = 5000;

                //Get Sequenced mapped Ports
                foreach (var mapItem in orderedEndCutPorts)
                {
                    edgeID = edgeID + 1;
                    sequencedMappedPorts.Add(edgeID, mapItem.Value);
                }

                OrderedEndCutPortsData = new OrderedEndCutPortsInformation(sequencedMappedPorts, orderedEndcutPortsInformation.EdgePoints, orderedEndcutPortsInformation.EdgeAngles,
                                                                            orderedEndcutPortsInformation.TopOrLeftInsidePort, orderedEndcutPortsInformation.TopOrLeftInsidePosition,
                                                                            orderedEndcutPortsInformation.BottomOrRightInsidePort, orderedEndcutPortsInformation.BottomOrRightInsidePosition);
            }
            else
            {
                throw new Exception("orderedEndCutPorts");
            }
            return OrderedEndCutPortsData;
        }

        #endregion

        /// <summary>
        /// Gets the assembly connection between bounded and bounding ports.
        /// </summary>
        /// <param name="boundingPort">The bounding port.</param>
        /// <param name="boundedPort">The bounded port.</param>
        /// <returns>Assembly Connection object.</returns>
        internal BusinessObject GetAssemblyConnection(IPort boundingPort, IPort boundedPort)
        {
            ReadOnlyCollection<IConnection> oConnectionCollection = null;
            Collection<AssemblyConnection> oAssemblyConnections = new Collection<AssemblyConnection>();
            boundedPort.Connectable.IsConnectedTo(boundingPort.Connectable, out oConnectionCollection);
            if (oConnectionCollection != null && oConnectionCollection.Count > 0)
            {
                foreach (IConnection oConnection in oConnectionCollection)
                {
                    if (oConnection is AssemblyConnection)
                    {
                        oAssemblyConnections.Add((AssemblyConnection)oConnection);
                    }
                }
                if (oAssemblyConnections != null)
                {
                    if (oAssemblyConnections.Count > 0)
                    {
                        // For now there are cases where only one object is filed into the collection. If any case has been encountered with more than one assembly connection, then get always the first object filled. 
                        AssemblyConnection assemblyConnection = oAssemblyConnections[0];
                        return assemblyConnection;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the primary assembly connection the bounded port is involved with; AND bounding port is auxiliary port.
        /// </summary>
        /// <param name="boundingPort">The bounding port.</param>
        /// <param name="boundedPort">The bounded port.</param>
        /// <returns>Assembly Connection object.</returns>
        internal BusinessObject GetPrimaryAssemblyConnection(IPort boundingPort, IPort boundedPort)
        {
            Collection<AssemblyConnection> oAssemblyConnections = new Collection<AssemblyConnection>();

            try
            {
                // Get all the Assembly Connections on BoundedPort
                ReadOnlyCollection<IConnection> connectionCollection = boundedPort.Connections;
                if (connectionCollection != null && connectionCollection.Count > 0)
                {
                    foreach (IConnection connection in connectionCollection)
                    {
                        if (connection is AssemblyConnection)
                        {
                            oAssemblyConnections.Add((AssemblyConnection)connection);
                        }
                    }
                    if (oAssemblyConnections != null)
                    {
                        // Iterate thru Assembly Connections and find the Assembly Conenction we are "interested" in..
                        // i.e this bounding port is part of auxiliary ports
                        for (int i = 0; i <= oAssemblyConnections.Count - 1; i++)
                        {
                            AssemblyConnection assemblyConnection = oAssemblyConnections[i];

                            if (assemblyConnection != null)
                            {

                                // "interested": Get the Auxiallary ports on this assembly connection and check if the input boundingPort is part of it
                                Collection<IPort> oAuxiliaryPorts = assemblyConnection.AuxiliaryPorts;
                                if (oAuxiliaryPorts != null)
                                {
                                    for (int j = 0; j <= oAuxiliaryPorts.Count - 1; j++)
                                    {
                                        if (oAuxiliaryPorts[j].Equals(boundingPort))
                                        {
                                            // YES, the boundingPort is Auxially Port.. Go ahead and pick this Assembly Connection
                                            return assemblyConnection;
                                        }
                                    }
                                }
                                return null;
                            }
                        }
                    }
                }
            }
            catch
            {
                throw new Exception("Failed to get primary Assembly connection");
            }

            return null;
        }


        //<summary>
        //Gets the cross section matrix of a stiffener at the specified position
        //</summary>
        //<param name="stiffener">The stiffener part.</param>
        //<param name="position">The position at which to retrieve the matrix.</param>
        internal bool HasAuxiliaryPorts(BusinessObject appConnection)
        {
            if (appConnection == null)
                return false;

            AssemblyConnection assemblyConnection = (AssemblyConnection)appConnection;
            Collection<IPort> oAuxiliaryPorts = null;
            if (assemblyConnection != null)
                oAuxiliaryPorts = assemblyConnection.AuxiliaryPorts;

            if (oAuxiliaryPorts != null)
                if (oAuxiliaryPorts.Count > 0)
                    return true;

            return false;
        }

        internal AssemblyConnection GetParentAssemblyConnection(BusinessObject ChildObject)
        {
            BusinessObject assemblyParent = SymbolHelper.GetCustomAssemblyParent(ChildObject);
            if (assemblyParent != null)
            {
                if (assemblyParent is AssemblyConnection)
                {
                    return (AssemblyConnection)assemblyParent;
                }
                else
                {
                    assemblyParent = GetParentAssemblyConnection(assemblyParent);
                }
            }

            return (AssemblyConnection)assemblyParent;

        }
    }
}

