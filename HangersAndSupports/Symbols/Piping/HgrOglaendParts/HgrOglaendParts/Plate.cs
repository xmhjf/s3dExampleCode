//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Plate.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.Plate
//   Author       :  Ramya	
//   Creation Date:  07.June.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.June.2012     VSP     Initial Creation
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
    public class Plate : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.Plate"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "FlangeWidth", "Width of Flange", 0)]
        public InputDouble m_FlangeWidth;
        [InputDouble(3, "WebWidth", "Width of Web", 0)]
        public InputDouble m_WebWidth;
        [InputDouble(4, "Length", "Length", 0.5)]
        public InputDouble m_Length;
        [InputDouble(5, "Thickness", "Thickness", 0)]
        public InputDouble m_Thickness;
        [InputDouble(6, "Offset1", "Offset1", 0)]
        public InputDouble m_Offset1;
        [InputDouble(7, "Offset2", "Offset2", 0)]
        public InputDouble m_Offset2;
        [InputDouble(8, "HP1PosX", "HPort1 xOffset of Plate", 0)]
        public InputDouble m_HP1xOffset;
        [InputDouble(9, "HP1PosY", "HPort1 yOffset of Plate", 0)]
        public InputDouble m_HP1yOffset;
        [InputDouble(10, "HP1PosZ", "HPort1 zOffset of Plate", 0)]
        public InputDouble m_HP1zOffset;
        [InputDouble(11, "HP2PosX", "HPort2 xOffset of Plate", 0)]
        public InputDouble m_HP2xOffset;
        [InputDouble(12, "HP2PosY", "HPort2 yOffset of Plate", 0)]
        public InputDouble m_HP2yOffset;
        [InputDouble(13, "HP2PosZ", "HPort2 zOffset of Plate", 0)]
        public InputDouble m_HP2zOffset;


        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Plate1", "Plate1")]
        [SymbolOutput("Plate2", "Plate2")]
        [SymbolOutput("Plate3", "Plate3")]
        [SymbolOutput("Plate4", "Plate4")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
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
                double flangeWidth = 0;
                double webWidth = 0;
                double thickness = 0;
                double offset1 = m_Offset1.Value;
                double offset2 = m_Offset2.Value;

                length = m_Length.Value;
                if (length < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Length value should be greater than 0"));
                    return;
                }

                flangeWidth = m_FlangeWidth.Value;
                if (flangeWidth < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Flange Width value should be greater than 0"));
                    return;
                }

                webWidth = m_WebWidth.Value;
                if (webWidth < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Web Width value should be greater than 0"));
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
                Line3d webLine1, webLine2, webLine3, webLine4, webLine5, webLine6, webLine7, webLine8;
                Collection<ICurve> webColl1 = new Collection<ICurve>();
                Collection<ICurve> webColl2 = new Collection<ICurve>();

                webLine1 = new Line3d(new Position(offset1 / 2, -(webWidth / 2), 0), new Position(offset1 / 2, -(webWidth / 2), length));
                webLine2 = new Line3d(new Position(offset1 / 2, -(webWidth / 2), length), new Position(offset1 / 2, (webWidth / 2), length));
                webLine3 = new Line3d(new Position(offset1 / 2, (webWidth / 2), length), new Position(offset1 / 2, (webWidth / 2), 0));
                webLine4 = new Line3d(new Position(offset1 / 2, (webWidth / 2), 0), new Position(offset1 / 2, -(webWidth / 2), 0));

                ////Add the curves into ICurves Collection
                webColl1.Add(webLine1);
                webColl1.Add(webLine2);
                webColl1.Add(webLine3);
                webColl1.Add(webLine4);

                webLine5 = new Line3d(new Position(-(offset1 / 2), -(webWidth / 2), 0), new Position(-(offset1 / 2), -(webWidth / 2), length));
                webLine6 = new Line3d(new Position(-(offset1 / 2), -(webWidth / 2), length), new Position(-(offset1 / 2), (webWidth / 2), length));
                webLine7 = new Line3d(new Position(-(offset1 / 2), (webWidth / 2), length), new Position(-(offset1 / 2), (webWidth / 2), 0));
                webLine8 = new Line3d(new Position(-(offset1 / 2), (webWidth / 2), 0), new Position(-(offset1 / 2), -(webWidth / 2), 0));

                ////Add the curves into ICurves Collection
                webColl2.Add(webLine5);
                webColl2.Add(webLine6);
                webColl2.Add(webLine7);
                webColl2.Add(webLine8);

                //Create the Flange
                Line3d flangeLine1, flangeLine2, flangeLine3, flangeLine4, flangeLine5, flangeLine6, flangeLine7, flangeLine8;
                Collection<ICurve> flangeColl1 = new Collection<ICurve>();
                Collection<ICurve> flangeColl2 = new Collection<ICurve>();

                flangeLine1 = new Line3d(new Position(-(flangeWidth / 2), (offset2 / 2), 0), new Position(-(flangeWidth / 2), (offset2 / 2), length));
                flangeLine2 = new Line3d(new Position(-(flangeWidth / 2), (offset2 / 2), length), new Position(flangeWidth / 2, (offset2 / 2), length));
                flangeLine3 = new Line3d(new Position(flangeWidth / 2, (offset2 / 2), length), new Position(flangeWidth / 2, (offset2 / 2), 0));
                flangeLine4 = new Line3d(new Position(flangeWidth / 2, (offset2 / 2), 0), new Position(-(flangeWidth / 2), (offset2 / 2), 0));

                flangeColl1.Add(flangeLine1);
                flangeColl1.Add(flangeLine2);
                flangeColl1.Add(flangeLine3);
                flangeColl1.Add(flangeLine4);

                flangeLine5 = new Line3d(new Position(-(flangeWidth / 2), -(offset2 / 2), 0), new Position(-(flangeWidth / 2), -(offset2 / 2), length));
                flangeLine6 = new Line3d(new Position(-(flangeWidth / 2), -(offset2 / 2), length), new Position(flangeWidth / 2, -(offset2 / 2), length));
                flangeLine7 = new Line3d(new Position(flangeWidth / 2, -(offset2 / 2), length), new Position(flangeWidth / 2, -(offset2 / 2), 0));
                flangeLine8 = new Line3d(new Position(flangeWidth / 2, -(offset2 / 2), 0), new Position(-(flangeWidth / 2), -(offset2 / 2), 0));

                flangeColl2.Add(flangeLine5);
                flangeColl2.Add(flangeLine6);
                flangeColl2.Add(flangeLine7);
                flangeColl2.Add(flangeLine8);

                ComplexString3d webComplxString1 = new ComplexString3d(webColl1);
                Vector webVector1 = new Vector(1, 0, 0);
                Projection3d webProjection1 = new Projection3d(connection, webComplxString1, webVector1, thickness, true);

                ComplexString3d flangeComplxString1 = new ComplexString3d(flangeColl1);
                Vector flangeVector1 = new Vector(0, 1, 0);
                Projection3d flangeProjection1 = new Projection3d(connection, flangeComplxString1, flangeVector1, thickness, true);

                ComplexString3d webComplxString2 = new ComplexString3d(webColl2);
                Vector webVector2 = new Vector(-1, 0, 0);
                Projection3d webProjection2 = new Projection3d(connection, webComplxString2, webVector2, thickness, true);

                ComplexString3d flangeComplxString2 = new ComplexString3d(flangeColl2);
                Vector flangeVector2 = new Vector(0, -1, 0);
                Projection3d flangeProjection2 = new Projection3d(connection, flangeComplxString2, flangeVector2, thickness, true);

                m_PhysicalAspect.Outputs["Plate1"] = webProjection1;
                m_PhysicalAspect.Outputs["Plate2"] = flangeProjection1;
                m_PhysicalAspect.Outputs["Plate3"] = webProjection2;
                m_PhysicalAspect.Outputs["Plate4"] = flangeProjection2;

                double hport1OffsetX = m_HP1xOffset.Value;
                double hport1OffsetY = m_HP1yOffset.Value;
                double hport1OffsetZ = m_HP1zOffset.Value;
                double hport2OffsetX = m_HP2xOffset.Value;
                double hport2OffsetY = m_HP2yOffset.Value;
                double hport2OffsetZ = m_HP2zOffset.Value;

                Port hole1Port = new Port(connection, part, "Hole1", new Position(0 + hport1OffsetX, 0 + hport1OffsetX, 0 + hport1OffsetX), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole1"] = hole1Port;

                Port hole2Port = new Port(connection, part, "Hole2", new Position(0 + hport1OffsetX, 0 + hport1OffsetX, length + hport1OffsetX), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole2"] = hole2Port;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    Type myType = this.GetType();
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Plate.cs " + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name));
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

