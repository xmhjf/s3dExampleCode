//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   CHgrPlate.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.CHgrPlate
//   Author       :  Ramya
//   Creation Date:  31.May.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31.May.2012     Ramya     Initial Creation
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
    public class CHgrPlate : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.CHgrPlate"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "FlangeWidth", "Width of Flange of L Plate", 0)]
        public InputDouble m_FlangeWidth;
        [InputDouble(3, "WebWidth", "Width of Web of L Plate", 0)]
        public InputDouble m_WebWidth;
        [InputDouble(4, "Length", "Length of L Plate", 0.5)]
        public InputDouble m_Length;
        [InputDouble(5, "OverhangFlWidth", "Overhang Flange Width of L Plate", 0)]
        public InputDouble m_OverhangFlWidth;
        [InputDouble(6, "OverhangWebWidth", "Overhang Web Width of L Plate", 0)]
        public InputDouble m_OverhangWebWidth;
        [InputDouble(7, "OverhangFlLength", "Overhang Flange Length of L Plate", 0)]
        public InputDouble m_OverhangFlLength;
        [InputDouble(8, "OverhangWebLength", "Overhang Web Length of L Plate", 0)]
        public InputDouble m_OverhangWebLength;
        [InputDouble(9, "Thickness", "Thickness of L Plate", 0)]
        public InputDouble m_Thickness;
        [InputDouble(10, "StuctureIntPosZ", "StuctureInt Port Position along Z", 0)]
        public InputDouble m_StuctureIntPosZ;
        [InputDouble(11, "Surface1PosZ", "Surface1 Port Position along Z", 0)]
        public InputDouble m_Surface1PosZ;
        [InputDouble(12, "Surface2PosZ", "Surface2 Port Position along Z", 0)]
        public InputDouble m_Surface2PosZ;
        [InputDouble(13, "Plate1Width", "Width of the plate1", 0, true)]
        public InputDouble m_Plate1Width;
        [InputDouble(14, "Plate1Depth", "Depth of the plate1", 0, true)]
        public InputDouble m_Plate1Depth;
        [InputDouble(15, "Plate1Thickness", "Thickness of the plate1", 0, true)]
        public InputDouble m_Plate1Thickness;
        [InputDouble(16, "Plate2Width", "Width of the plate2", 0, true)]
        public InputDouble m_Plate2Width;
        [InputDouble(17, "Plate2Depth", "Depth of the plate2", 0, true)]
        public InputDouble m_Plate2Depth;
        [InputDouble(18, "Plate2Thickness", "Thickness of the plate2", 0, true)]
        public InputDouble m_Plate2Thickness;
        [InputDouble(19, "HP1PosX", "HPort1 xOffset of Plate", 0)]
        public InputDouble m_HP1xOffset;
        [InputDouble(20, "HP1PosY", "HPort1 yOffset of Plate", 0)]
        public InputDouble m_HP1yOffset;
        [InputDouble(21, "HP1PosZ", "HPort1 zOffset of Plate", 0)]
        public InputDouble m_HP1zOffset;
        [InputDouble(22, "HP2PosX", "HPort2 xOffset of Plate", 0)]
        public InputDouble m_HP2xOffset;
        [InputDouble(23, "HP2PosY", "HPort2 yOffset of Plate", 0)]
        public InputDouble m_HP2yOffset;
        [InputDouble(24, "HP2PosZ", "HPort2 zOffset of Plate", 0)]
        public InputDouble m_HP2zOffset;


        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Flange1", "Flange1")]
        [SymbolOutput("Web", "Web")]
        [SymbolOutput("Flange2", "Flange2")]
        [SymbolOutput("Plate1", "Plate1")]
        [SymbolOutput("Plate2", "Plate2")]
        [SymbolOutput("Structure", "Structure")]
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
                double overhangFlWidth = m_OverhangFlWidth.Value;
                double overhangWebWidth = m_OverhangWebWidth.Value;
                double overhangFlLength = m_OverhangFlLength.Value;
                double overhangWebLength = m_OverhangWebLength.Value;
                double stuctureIntPosZ = m_StuctureIntPosZ.Value;
                double surface1PosZ = m_Surface1PosZ.Value;
                double surface2PosZ = m_Surface2PosZ.Value;
                double plate1Width = m_Plate1Width.Value;
                double plate1Depth = m_Plate1Depth.Value;
                double plate1Thickness = m_Plate1Thickness.Value;
                double plate2Width = m_Plate2Width.Value;
                double plate2Depth = m_Plate2Depth.Value;
                double plate2Thickness = m_Plate2Thickness.Value;

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

                //Create the Flange
                Line3d flangeLine1, flangeLine2, flangeLine3, flangeLine4, flangeLine5, flangeLine6;
                Collection<ICurve> flangeColl = new Collection<ICurve>();
                flangeLine1 = new Line3d(new Position(thickness, 0, 0), new Position(thickness, 0, length));

                if (overhangFlWidth > 0 && overhangFlLength > 0)
                {
                    //Create the Lines 
                    flangeLine2 = new Line3d(new Position(thickness, 0, length), new Position((flangeWidth - overhangFlWidth), 0, length));
                    flangeLine3 = new Line3d(new Position((flangeWidth - overhangFlWidth), 0, length), new Position((flangeWidth - overhangFlWidth), 0, (length + overhangFlLength)));
                    flangeLine4 = new Line3d(new Position((flangeWidth - overhangFlWidth), 0, (length + overhangFlLength)), new Position((flangeWidth), 0, (length + overhangFlLength)));
                    flangeLine5 = new Line3d(new Position((flangeWidth), 0, (length + overhangFlLength)), new Position(flangeWidth, 0, 0));
                    flangeLine6 = new Line3d(new Position(flangeWidth, 0, 0), new Position(thickness, 0, 0));

                    //Add the curves into ICurves Collection
                    flangeColl.Add(flangeLine1);
                    flangeColl.Add(flangeLine2);
                    flangeColl.Add(flangeLine3);
                    flangeColl.Add(flangeLine4);
                    flangeColl.Add(flangeLine5);
                    flangeColl.Add(flangeLine6);
                }
                else
                {
                    flangeLine2 = new Line3d(new Position(thickness, 0, length), new Position(flangeWidth, 0, length));
                    flangeLine3 = new Line3d(new Position(flangeWidth, 0, length), new Position(flangeWidth, 0, 0));
                    flangeLine4 = new Line3d(new Position(flangeWidth, 0, 0), new Position(thickness, 0, 0));

                    //Add the curves into ICurves Collection
                    flangeColl.Add(flangeLine1);
                    flangeColl.Add(flangeLine2);
                    flangeColl.Add(flangeLine3);
                    flangeColl.Add(flangeLine4);
                }

                //Create the Web
                Line3d webLine1, webLine2, webLine3, webLine4, webLine5, webLine6, webLine7, webLine8;
                Collection<ICurve> webColl = new Collection<ICurve>();

                if (overhangWebWidth > 0 && overhangWebLength > 0)
                {
                    webLine1 = new Line3d(new Position(0, 0, 0), new Position(0, 0, length));
                    webLine2 = new Line3d(new Position(0, 0, length), new Position(0, (webWidth - overhangWebWidth) / 2, length));
                    webLine3 = new Line3d(new Position(0, (webWidth - overhangWebWidth) / 2, length), new Position(0, (webWidth - overhangWebWidth) / 2, length + overhangWebLength));
                    webLine4 = new Line3d(new Position(0, (webWidth - overhangWebWidth) / 2, length + overhangWebLength), new Position(0, overhangWebWidth + (webWidth - overhangWebWidth) / 2, length + overhangWebLength));
                    webLine5 = new Line3d(new Position(0, overhangWebWidth + (webWidth - overhangWebWidth) / 2, length + overhangWebLength), new Position(0, overhangWebWidth + (webWidth - overhangWebWidth) / 2, length));
                    webLine6 = new Line3d(new Position(0, overhangWebWidth + (webWidth - overhangWebWidth) / 2, length), new Position(0, webWidth, length));
                    webLine7 = new Line3d(new Position(0, webWidth, length), new Position(0, webWidth, 0));
                    webLine8 = new Line3d(new Position(0, webWidth, 0), new Position(0, 0, 0));
                    ////Add the curves into ICurves Collection
                    webColl.Add(webLine1);
                    webColl.Add(webLine2);
                    webColl.Add(webLine3);
                    webColl.Add(webLine4);
                    webColl.Add(webLine5);
                    webColl.Add(webLine6);
                    webColl.Add(webLine7);
                    webColl.Add(webLine8);
                }
                else
                {
                    webLine1 = new Line3d(new Position(0, 0, 0), new Position(0, 0, length));
                    webLine2 = new Line3d(new Position(0, 0, length), new Position(0, webWidth, length));
                    webLine3 = new Line3d(new Position(0, webWidth, length), new Position(0, webWidth, 0));
                    webLine4 = new Line3d(new Position(0, webWidth, 0), new Position(0, 0, 0));

                    ////Add the curves into ICurves Collection
                    webColl.Add(webLine1);
                    webColl.Add(webLine2);
                    webColl.Add(webLine3);
                    webColl.Add(webLine4);
                }

                //Create the Flange
                Line3d flangeLine11, flangeLine12, flangeLine13, flangeLine14, flangeLine15, flangeLine16;
                Collection<ICurve> flangeColl1 = new Collection<ICurve>();
                flangeLine11 = new Line3d(new Position(thickness, webWidth - thickness, 0), new Position(thickness, webWidth - thickness, length));

                if (overhangFlWidth > 0 && overhangFlLength > 0)
                {
                    //Create the Lines 
                    flangeLine12 = new Line3d(new Position(thickness, webWidth - thickness, length), new Position((flangeWidth - overhangFlWidth), webWidth - thickness, length));
                    flangeLine13 = new Line3d(new Position((flangeWidth - overhangFlWidth), webWidth - thickness, length), new Position((flangeWidth - overhangFlWidth), webWidth - thickness, (length + overhangFlLength)));
                    flangeLine14 = new Line3d(new Position((flangeWidth - overhangFlWidth), webWidth - thickness, (length + overhangFlLength)), new Position((flangeWidth), webWidth - thickness, (length + overhangFlLength)));
                    flangeLine15 = new Line3d(new Position((flangeWidth), webWidth - thickness, (length + overhangFlLength)), new Position(flangeWidth, webWidth - thickness, 0));
                    flangeLine16 = new Line3d(new Position(flangeWidth, webWidth - thickness, 0), new Position(thickness, webWidth - thickness, 0));

                    //Add the curves into ICurves Collection
                    flangeColl1.Add(flangeLine11);
                    flangeColl1.Add(flangeLine12);
                    flangeColl1.Add(flangeLine13);
                    flangeColl1.Add(flangeLine14);
                    flangeColl1.Add(flangeLine15);
                    flangeColl1.Add(flangeLine16);
                }
                else
                {
                    flangeLine12 = new Line3d(new Position(thickness, webWidth - thickness, length), new Position(flangeWidth, webWidth - thickness, length));
                    flangeLine13 = new Line3d(new Position(flangeWidth, webWidth - thickness, length), new Position(flangeWidth, webWidth - thickness, 0));
                    flangeLine14 = new Line3d(new Position(flangeWidth, webWidth - thickness, 0), new Position(thickness, webWidth - thickness, 0));

                    //Add the curves into ICurves Collection
                    flangeColl1.Add(flangeLine11);
                    flangeColl1.Add(flangeLine12);
                    flangeColl1.Add(flangeLine13);
                    flangeColl1.Add(flangeLine14);
                }

                ComplexString3d webComplxString = new ComplexString3d(webColl);
                Vector webVector = new Vector(1, 0, 0);
                Projection3d webProjection = new Projection3d(connection, webComplxString, webVector, thickness, true);

                ComplexString3d flangeComplxString = new ComplexString3d(flangeColl);
                Vector flangeVector = new Vector(0, 1, 0);
                Projection3d flangeProjection = new Projection3d(connection, flangeComplxString, flangeVector, thickness, true);

                ComplexString3d flangeComplxString1 = new ComplexString3d(flangeColl1);
                Vector flangeVector1 = new Vector(0, 1, 0);
                Projection3d flangeProjection1 = new Projection3d(connection, flangeComplxString1, flangeVector1, thickness, true);


                m_PhysicalAspect.Outputs["Web"] = webProjection;
                m_PhysicalAspect.Outputs["Flange1"] = flangeProjection;
                m_PhysicalAspect.Outputs["Flange2"] = flangeProjection1;

                if (plate1Width > 0 && plate1Depth > 0 && plate1Thickness > 0)
                {
                    symbolGeomHlpr.ActivePosition = new Position(flangeWidth / 2, webWidth / 2, length);
                    m_PhysicalAspect.Outputs["Plate1"] = symbolGeomHlpr.CreateBox(connection, plate1Thickness, plate1Depth, plate1Width);
                }
                if (plate2Width > 0 && plate2Depth > 0 && plate2Thickness > 0)
                {
                    symbolGeomHlpr.ActivePosition = new Position(flangeWidth / 2, webWidth / 2, length + plate1Thickness);
                    m_PhysicalAspect.Outputs["Plate2"] = symbolGeomHlpr.CreateBox(connection, plate2Thickness, plate2Depth, plate2Width);
                }

                double hport1OffsetX = m_HP1xOffset.Value;
                double hport1OffsetY = m_HP1yOffset.Value;
                double hport1OffsetZ = m_HP1zOffset.Value;
                double hport2OffsetX = m_HP2xOffset.Value;
                double hport2OffsetY = m_HP2yOffset.Value;
                double hport2OffsetZ = m_HP2zOffset.Value;
                Port structurePort = null;

                if (overhangWebLength > 0)
                {
                    structurePort = new Port(connection, part, "Structure", new Position(((flangeWidth) / 2), (webWidth / 2), length + overhangWebLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Structure"] = structurePort;
                }
                else if (overhangFlLength > 0)
                {
                    structurePort = new Port(connection, part, "Structure", new Position(((flangeWidth) / 2), (webWidth / 2), length + overhangFlLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Structure"] = structurePort;
                }
                else
                {
                    structurePort = new Port(connection, part, "Structure", new Position((flangeWidth / 2), (webWidth / 2), length + plate1Thickness + plate2Thickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Structure"] = structurePort;
                }

                Port hole1Port = new Port(connection, part, "Hole1", new Position((thickness + (flangeWidth - thickness) / 2) + hport1OffsetX, thickness + hport1OffsetY, length - stuctureIntPosZ + hport1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole1"] = hole1Port;

                Port hole2Port = new Port(connection, part, "Hole2", new Position((thickness + (flangeWidth - thickness) / 2) + hport2OffsetX, webWidth - thickness + hport2OffsetY, length - stuctureIntPosZ + hport2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Hole2"] = hole2Port;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    Type myType = this.GetType();
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CHgrPlate.cs " + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name));
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