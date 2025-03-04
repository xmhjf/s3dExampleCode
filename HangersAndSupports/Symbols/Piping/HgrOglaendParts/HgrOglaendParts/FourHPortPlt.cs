//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   FourHPortPlt.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.FourHPortPlt
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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class FourHPortPlt : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HS_OglaendParts,Ingr.SP3D.Content.Support.Symbols.FourHPortPlt"
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
        [InputDouble(5, "HP1PosX", "HPort1 xOffset of Plate", 0)]
        public InputDouble m_HP1xOffset;
        [InputDouble(6, "HP1PosY", "HPort1 yOffset of Plate", 0)]
        public InputDouble m_HP1yOffset;
        [InputDouble(7, "HP1PosZ", "HPort1 zOffset of Plate", 0)]
        public InputDouble m_HP1zOffset;
        [InputDouble(8, "HP2PosX", "HPort2 xOffset of Plate", 0)]
        public InputDouble m_HP2xOffset;
        [InputDouble(9, "HP2PosY", "HPort2 yOffset of Plate", 0)]
        public InputDouble m_HP2yOffset;
        [InputDouble(10, "HP2PosZ", "HPort2 zOffset of Plate", 0)]
        public InputDouble m_HP2zOffset;
        [InputDouble(11, "HP3PosX", "HPort3 xOffset of Plate", 0)]
        public InputDouble m_HP3xOffset;
        [InputDouble(12, "HP3PosY", "HPort3 yOffset of Plate", 0)]
        public InputDouble m_HP3yOffset;
        [InputDouble(13, "HP3PosZ", "HPort3 zOffset of Plate", 0)]
        public InputDouble m_HP3zOffset;
        [InputDouble(14, "HP4PosX", "HPort4 xOffset of Plate", 0)]
        public InputDouble m_HP4xOffset;
        [InputDouble(15, "HP4PosY", "HPort4 yOffset of Plate", 0)]
        public InputDouble m_HP4yOffset;
        [InputDouble(16, "HP4PosZ", "HPort4 zOffset of Plate", 0)]
        public InputDouble m_HP4zOffset;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Plate", "Plate")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
        [SymbolOutput("Hole3", "Hole3")]
        [SymbolOutput("Hole4", "Hole4")]

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
                //Collection<ICurve> webColl1 = new Collection<ICurve>();
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

                double hport1OffsetX = m_HP1xOffset.Value;
                double hport1OffsetY = m_HP1yOffset.Value;
                double hport1OffsetZ = m_HP1zOffset.Value;
                double hport2OffsetX = m_HP2xOffset.Value;
                double hport2OffsetY = m_HP2yOffset.Value;
                double hport2OffsetZ = m_HP2zOffset.Value;
                double hport3OffsetX = m_HP3xOffset.Value;
                double hport3OffsetY = m_HP3yOffset.Value;
                double hport3OffsetZ = m_HP3zOffset.Value;
                double hport4OffsetX = m_HP4xOffset.Value;
                double hport4OffsetY = m_HP4yOffset.Value;
                double hport4OffsetZ = m_HP4zOffset.Value;

                m_PhysicalAspect.Outputs["Plate"] = plateProjection;


                Port hole1Port = new Port(connection, part, "Hole1", new Position(width / 2 + hport1OffsetX, 0 + hport1OffsetY, length / 2 - 0.005 + hport1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole1"] = hole1Port;

                Port hole2Port = new Port(connection, part, "Hole2", new Position(width / 2 + hport2OffsetX, 0 + hport2OffsetY, length / 2 + 0.005 + hport2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole2"] = hole2Port;

                Port hole3Port = new Port(connection, part, "Hole3", new Position(width / 2 + hport3OffsetX, thickness + hport3OffsetY, length / 2 - 0.005 + hport3OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole3"] = hole3Port;

                Port hole4Port = new Port(connection, part, "Hole4", new Position(width / 2 + hport4OffsetX, thickness + hport4OffsetY, length / 2 + 0.005 + hport4OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole4"] = hole4Port;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    Type myType = this.GetType();
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of FourHPortPlt.cs " + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name));
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