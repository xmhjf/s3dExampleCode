//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   TwoPortPlt.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.TwoPortPlt.cs
//   Author       :  PVS	
//   Creation Date:  07.June.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.June.2012     PVS     Initial Creation
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.Support.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class TwoPortPlt : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.TwoPortPlt"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width1", "Width of Plate", 0)]
        public InputDouble m_Width;
        [InputDouble(3, "Length1", "Length of Plate", 0)]
        public InputDouble m_Length;
        [InputDouble(4, "Thickness1", "Thickness of Plate", 0)]
        public InputDouble m_Thickness;
        [InputDouble(5, "P1xOffset", "Port1 xOffset of Plate", 0)]
        public InputDouble m_P1xOffset;
        [InputDouble(6, "P1yOffset", "Port1 yOffset of Plate", 0)]
        public InputDouble m_P1yOffset;
        [InputDouble(7, "P1zOffset", "Port1 zOffset of Plate", 0)]
        public InputDouble m_P1zOffset;
        [InputDouble(8, "P2xOffset", "Port2 xOffset of Plate", 0)]
        public InputDouble m_P2xOffset;
        [InputDouble(9, "P2yOffset", "Port2 yOffset of Plate", 0)]
        public InputDouble m_P2yOffset;
        [InputDouble(10, "P2zOffset", "Port2 zOffset of Plate", 0)]
        public InputDouble m_P2zOffset;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Plate", "Plate")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = null;
                SymbolGeometryHelper symbolGeomHlpr = new SymbolGeometryHelper();
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                part = m_PartInput.Value as Part;
                double length = 0;
                double width = 0;
                double thickness = 0;

                length = m_Length.Value;
                if (length < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Length value should be greater than 0"));
                    return;
                }

                width = m_Width.Value;
                if (width < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Width value should be greater than 0"));
                    return;
                }

                thickness = m_Thickness.Value;
                if (thickness < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Thickness value should be greater than 0"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //Initialize SymbolGeometryHelper. Set the active position and orientation 
                symbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                symbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                //Create the Web
                Line3d plateLine1, plateLine2, plateLine3, plateLine4;
                Collection<ICurve> plateColl1 = new Collection<ICurve>();

                plateLine1 = new Line3d(new Position(0, 0, 0), new Position(0, 0, length));
                plateLine2 = new Line3d(new Position(0, 0, length), new Position(width, 0, length));
                plateLine3 = new Line3d(new Position(width, 0, length), new Position(width, 0, 0));
                plateLine4 = new Line3d(new Position(width, 0, 0), new Position(0, 0, 0));

                ////Add the curves into ICurves Collection
                plateColl1.Add(plateLine1);
                plateColl1.Add(plateLine2);
                plateColl1.Add(plateLine3);
                plateColl1.Add(plateLine4);

                ComplexString3d plateComplxString = new ComplexString3d(plateColl1);
                Vector plateVector = new Vector(0, 1, 0);
                Projection3d plateProjection = new Projection3d(connection, plateComplxString, plateVector, thickness, true);

                double port1OffsetX = m_P1xOffset.Value;
                double port1OffsetY = m_P1yOffset.Value;
                double port1OffsetZ = m_P1zOffset.Value;
                double port2OffsetX = m_P2xOffset.Value;
                double port2OffsetY = m_P2yOffset.Value;
                double port2OffsetZ = m_P2zOffset.Value;

                m_PhysicalAspect.Outputs["Plate"] = plateProjection;

                Port port1 = new Port(connection, part, "Port1", new Position(width / 2 - port1OffsetX, 0 - port1OffsetY, length - port1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "Port2", new Position(width / 2 - port2OffsetX, thickness - port2OffsetY, length - port2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = port2;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    Type myType = this.GetType();
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of TwoPortPlt.cs " + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name));
                }
            }

        #endregion
        }
#region ICustomHgrWeightCG Members
        void ICustomHgrWeightCG.WeightCG(SupportComponent supportComponent, ref double weight, ref double cogX, ref double cogY, ref double cogZ)
        {

            ////System WCG Attributes

            Part catalogPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

            try
            {
                weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
            }
            catch
            {
                weight = 0;
            }
            //Center of Gravity
            try
            {
                cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
            }
            catch
            {
                cogX = 0;
            }
            try
            {
                cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
            }
            catch
            {
                cogY = 0;
            }
            try
            {
                cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
            }
            catch
            {
                cogZ = 0;
            }

        }
        #endregion
    }
}

