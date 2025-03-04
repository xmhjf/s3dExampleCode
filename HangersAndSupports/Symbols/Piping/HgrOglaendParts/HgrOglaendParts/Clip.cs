//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Clip.cs
//   HS_OglaendParts,Ingr.SP3D.Content.Support.Symbols.Clip
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
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Collections.ObjectModel;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [VariableOutputs]
    public class Clip : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HS_OglaendParts,Ingr.SP3D.Content.Support.Symbols.Clip"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)] public InputCatalogPart m_oPartInput;
        [InputDouble(2, "ClipPlt1Width", "Width of the ClipPlate1", 0, true)] public InputDouble m_dClipPlt1Width;
        [InputDouble(3, "ClipPlt1Depth", "Depth of the ClipPlate1", 0, true)] public InputDouble m_dClipPlt1Depth;
        [InputDouble(4, "ClipPlt1Thickness", "Thickness of the ClipPlate1", 0, true)] public InputDouble m_dClipPlt1Thickness;
        [InputDouble(5, "ClipPlt2Width", "Width of the ClipPlate2", 0, true)] public InputDouble m_dClipPlt2Width;
        [InputDouble(6, "ClipPlt2Depth", "Depth of the ClipPlate2", 0, true)] public InputDouble m_dClipPlt2Depth;
        [InputDouble(7, "ClipPlt2Thickness", "Thickness of the ClipPlate2", 0, true)] public InputDouble m_dClipPlt2Thickness;
        [InputDouble(8, "HDistance", "Horizontal Distance", 0, true)] public InputDouble m_dHorDistance;
        [InputDouble(9, "VDistance", "Vertical Distance", 0, true)] public InputDouble m_dVerDistance;
        [InputDouble(10, "ClipGap", "Gap Between Clips", 0, true)] public InputDouble m_dClipGap;
        [InputDouble(11, "PlateOffSet", "OffSet Between Clips and Plates", 0, true)]public InputDouble m_dPlateOff;
        [InputDouble(12, "Plate1Width", "Depth of the plate1", 0, true)] public InputDouble m_dPlate1Width;
        [InputDouble(13, "Plate1Depth", "Depth of the plate1", 0, true)] public InputDouble m_dPlate1Depth;
        [InputDouble(14, "Plate1Thickness", "Thickness of the plate1", 0, true)] public InputDouble m_dPlate1Thickness;
        [InputDouble(15, "Plate2Width", "Width of the plate2", 0, true)] public InputDouble m_dPlate2Width;
        [InputDouble(16, "Plate2Depth", "Depth of the plate2", 0, true)] public InputDouble m_dPlate2Depth;
        [InputDouble(17, "Plate2Thickness", "Thickness of the plate2", 0, true)] public InputDouble m_dPlate2Thickness;
        [InputDouble(18, "OHP1PosX", "HPort1 xOffset of Plate", 0)] public InputDouble m_dHP1xOffset;
        [InputDouble(19, "OHP1PosY", "HPort1 yOffset of Plate", 0)] public InputDouble m_dHP1yOffset;
        [InputDouble(20, "OHP1PosZ", "HPort1 zOffset of Plate", 0)] public InputDouble m_dHP1zOffset;
        [InputDouble(21, "OHP2PosX", "HPort2 xOffset of Plate", 0)] public InputDouble m_dHP2xOffset;
        [InputDouble(22, "OHP2PosY", "HPort2 yOffset of Plate", 0)] public InputDouble m_dHP2yOffset;
        [InputDouble(23, "OHP2PosZ", "HPort2 zOffset of Plate", 0)] public InputDouble m_dHP2zOffset;
        [InputDouble(24, "SPPosX", "SteePort xOffset of Plate", 0)] public InputDouble m_dSPxOffset;
        [InputDouble(25, "SPPosY", "SteePort yOffset of Plate", 0)] public InputDouble m_dSPyOffset;
        [InputDouble(26, "SPPosZ", "SteePort zOffset of Plate", 0)] public InputDouble m_dSPzOffset;
        [InputDouble(27, "PlateHorOffSet", "Horizontal OffSet Between Clips and Plates", 0, true)]public InputDouble m_dPlateHorOff;
        [InputDouble(28, "PlateWidth", "Depth of the plate", 0, true)]public InputDouble m_dPlateWidth;
        [InputDouble(29, "PlateDepth", "Depth of the plate", 0, true)]public InputDouble m_dPlateDepth;
        [InputDouble(30, "PlateThickness", "Thickness of the plate", 0, true)]public InputDouble m_dPlateThickness;
        [InputDouble(31, "ScrewLength", "Depth of the plate", 0, true)]        public InputDouble m_dScrewLength;
        [InputDouble(32, "ScrewDia", "Thickness of the plate", 0, true)]        public InputDouble m_dScrewDia;
        [InputDouble(33, "ScrewVOffset", "Screw Vertical Offset", 0, true)]        public InputDouble m_dScrewVOffset;
        [InputDouble(34, "ScrewHOffset1", "Screw Vertical Offset", 0, true)]        public InputDouble m_dScrewHOffset1;
        [InputDouble(35, "ScrewHOffset2", "Screw Vertical Offset", 0, true)]        public InputDouble m_dScrewHOffset2;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Clip1", "Clip1")]
        [SymbolOutput("Clip2", "Clip2")]
        [SymbolOutput("Plate1", "Plate1")]
        [SymbolOutput("Plate2", "Plate2")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
        [SymbolOutput("Steel", "Steel")]
        [SymbolOutput("Screw1", "Screw")]
        [SymbolOutput("Screw2", "Screw")]
        
        public AspectDefinition m_oPhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part oPart = null;
                SymbolGeometryHelper oSymbolGeomHlpr = new SymbolGeometryHelper();
                SP3DConnection oConnection = default(SP3DConnection);
                oConnection = OccurrenceConnection;

                oPart = m_oPartInput.Value as Part;

                double dClipPlt1Width = m_dClipPlt1Width.Value;
                double dClipPlt1Depth = m_dClipPlt1Depth.Value;
                double dClipPlt1Thickness = m_dClipPlt1Thickness.Value;
                double dClipPlt2Width = m_dClipPlt2Width.Value;
                double dClipPlt2Depth = m_dClipPlt2Depth.Value;
                double dClipPlt2Thickness = m_dClipPlt2Thickness.Value;
                double dPlateWidth = m_dPlateWidth.Value;
                double dPlateDepth = m_dPlateDepth.Value;
                double dPlateThickness = m_dPlateThickness.Value;
                double dPlate1Width = m_dPlate1Width.Value;
                double dPlate1Depth = m_dPlate1Depth.Value;
                double dPlate1Thickness = m_dPlate1Thickness.Value;
                double dPlate2Width = m_dPlate2Width.Value;
                double dPlate2Depth = m_dPlate2Depth.Value;
                double dPlate2Thickness = m_dPlate2Thickness.Value;
                double dHorDistance = m_dHorDistance.Value;
                double dVerDistance = m_dVerDistance.Value;
                double dClipGap = m_dClipGap.Value;
                double dPlateOff = m_dPlateOff.Value;
                double dPlateHorOff = m_dPlateHorOff.Value;
                double dScrewLength = m_dScrewLength.Value;
                double dScrewDia = m_dScrewDia.Value;
                double dScrewHOffset1 = m_dScrewHOffset1.Value;
                double dScrewHOffset2 = m_dScrewHOffset2.Value;
                double dScrewVOffset = m_dScrewVOffset.Value;


                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //Initialize SymbolGeometryHelper. Set the active position and orientation 

                oSymbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                oSymbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                //Create the Flange
                ////Plate1 Creation

                Line3d oClip11, oClip12, oClip13, oClip14, oClip15, oClip16, oClip17, oClip18;
                Collection<ICurve> oClip1Coll = new Collection<ICurve>();

                oClip11 = new Line3d(new Position(0, 0, 0), new Position(0, dClipPlt1Width, 0));
                oClip12 = new Line3d(new Position(0, dClipPlt1Width, 0), new Position(0, dHorDistance + dClipPlt1Width, dVerDistance));
                oClip13 = new Line3d(new Position(0, dHorDistance + dClipPlt1Width, dVerDistance), new Position(0, dClipPlt1Width + dHorDistance + dClipPlt2Width, dVerDistance));
                oClip14 = new Line3d(new Position(0, dClipPlt1Width + dHorDistance + dClipPlt2Width, dVerDistance), new Position(0, dClipPlt1Width + dHorDistance + dClipPlt2Width, dVerDistance + dClipPlt2Thickness));
                oClip15 = new Line3d(new Position(0, dClipPlt1Width + dHorDistance + dClipPlt2Width, dVerDistance + dClipPlt2Thickness), new Position(0, dClipPlt1Width + dHorDistance, dVerDistance + dClipPlt2Thickness));
                oClip16 = new Line3d(new Position(0, dClipPlt1Width + dHorDistance, dVerDistance + dClipPlt2Thickness), new Position(0, dClipPlt1Width, dClipPlt1Thickness));
                oClip17 = new Line3d(new Position(0, dClipPlt1Width, dClipPlt1Thickness), new Position(0, 0, dClipPlt1Thickness));
                oClip18 = new Line3d(new Position(0, 0, dClipPlt1Thickness), new Position(0, 0, 0));

                //dPlate1Thickness
                oClip1Coll.Add(oClip11);
                oClip1Coll.Add(oClip12);
                oClip1Coll.Add(oClip13);
                oClip1Coll.Add(oClip14);
                oClip1Coll.Add(oClip15);
                oClip1Coll.Add(oClip16);
                oClip1Coll.Add(oClip17);
                oClip1Coll.Add(oClip18);

                ComplexString3d oClip1ComplxString = new ComplexString3d(oClip1Coll);
                Vector Clip1Vector = new Vector(1, 0, 0);
                Projection3d oClip1Projection = new Projection3d(oConnection, oClip1ComplxString, Clip1Vector, dClipPlt1Depth, true);

                m_oPhysicalAspect.Outputs["Clip1"] = oClip1Projection;

                //Add Screw

                oSymbolGeomHlpr = new SymbolGeometryHelper();
                oSymbolGeomHlpr.ActivePosition = new Position(dScrewHOffset1, dClipPlt1Width / 2, -(dScrewLength / 2 - dClipPlt1Thickness) + dScrewVOffset);
                oSymbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d cylinder1 = (Projection3d)oSymbolGeomHlpr.CreateCylinder(OccurrenceConnection, dScrewDia / 2, dScrewLength);
                m_oPhysicalAspect.Outputs["Screw1"] = cylinder1;

                oSymbolGeomHlpr = new SymbolGeometryHelper();
                oSymbolGeomHlpr.ActivePosition = new Position(dScrewHOffset2, dClipPlt1Width / 2, -(dScrewLength / 2 - dClipPlt1Thickness) + dScrewVOffset);
                oSymbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d cylinder2 = (Projection3d)oSymbolGeomHlpr.CreateCylinder(OccurrenceConnection, dScrewDia / 2, dScrewLength);
                m_oPhysicalAspect.Outputs["Screw2"] = cylinder2;




                if (dPlate1Width > 0 && dPlate1Depth > 0 && dPlate1Thickness > 0)
                {
                    if (dPlate2Width > 0 && dPlate2Depth > 0 && dPlate2Thickness > 0)
                    {
                        Line3d oPlate21, oPlate22, oPlate23, oPlate24;
                        Collection<ICurve> oPlateColl = new Collection<ICurve>();

                        oPlate21 = new Line3d(new Position(0, -dPlateHorOff, -dPlateOff), new Position(0, -dPlateHorOff + dPlate2Width, -dPlateOff));
                        oPlate22 = new Line3d(new Position(0, -dPlateHorOff + dPlate2Width, -dPlateOff), new Position(0, dPlate2Width - dPlateHorOff, -dPlate2Thickness - dPlateOff));
                        oPlate23 = new Line3d(new Position(0, -dPlateHorOff + dPlate2Width, -dPlate2Thickness - dPlateOff), new Position(0, -dPlateHorOff, -dPlate2Thickness - dPlateOff));
                        oPlate24 = new Line3d(new Position(0, -dPlateHorOff, -dPlate2Thickness - dPlateOff), new Position(0, -dPlateHorOff, -dPlateOff));


                        oPlateColl.Add(oPlate21);
                        oPlateColl.Add(oPlate22);
                        oPlateColl.Add(oPlate23);
                        oPlateColl.Add(oPlate24);

                        ComplexString3d oPlateComplxString = new ComplexString3d(oPlateColl);
                        Vector PlateVector = new Vector(1, 0, 0);
                        Projection3d oPlateProjection = new Projection3d(oConnection, oPlateComplxString, PlateVector, dPlate2Depth, true);

                        m_oPhysicalAspect.Outputs["Plate2"] = oPlateProjection;


                        Line3d oPlate11, oPlate12, oPlate13, oPlate14;
                        Collection<ICurve> oPlate1Coll = new Collection<ICurve>();

                        oPlate11 = new Line3d(new Position(0, 0 - dPlateHorOff, -dPlateOff - dPlate2Thickness), new Position(0, 0 - dPlateHorOff + dPlate1Width, -dPlateOff - dPlate2Thickness));
                        oPlate12 = new Line3d(new Position(0, 0 - dPlateHorOff + dPlate1Width, -dPlateOff - dPlate2Thickness), new Position(0, 0 + dPlate1Width - dPlateHorOff, -dPlate1Thickness - dPlateOff - dPlate2Thickness));
                        oPlate13 = new Line3d(new Position(0, 0 - dPlateHorOff + dPlate1Width, -dPlate1Thickness - dPlateOff - dPlate2Thickness), new Position(0, 0 - dPlateHorOff, -dPlate1Thickness - dPlateOff - dPlate2Thickness));
                        oPlate14 = new Line3d(new Position(0, 0 - dPlateHorOff, -dPlate1Thickness - dPlateOff - dPlate2Thickness), new Position(0, 0 - dPlateHorOff, -dPlateOff - dPlate2Thickness));


                        oPlate1Coll.Add(oPlate11);
                        oPlate1Coll.Add(oPlate12);
                        oPlate1Coll.Add(oPlate13);
                        oPlate1Coll.Add(oPlate14);

                        ComplexString3d oPlate1ComplxString = new ComplexString3d(oPlate1Coll);
                        Vector Plate1Vector = new Vector(1, 0, 0);
                        Projection3d oPlate1Projection = new Projection3d(oConnection, oPlate1ComplxString, Plate1Vector, dPlate1Depth, true);

                        m_oPhysicalAspect.Outputs["Plate1"] = oPlate1Projection;
                    }
                    else
                    {
                        Line3d oPlate11, oPlate12, oPlate13, oPlate14;
                        Collection<ICurve> oPlate1Coll = new Collection<ICurve>();

                        oPlate11 = new Line3d(new Position(0, 0 - dPlateHorOff, -dPlateOff), new Position(0, 0 - dPlateHorOff + dPlate1Width, -dPlateOff));
                        oPlate12 = new Line3d(new Position(0, 0 - dPlateHorOff + dPlate1Width, -dPlateOff), new Position(0, 0 + dPlate1Width - dPlateHorOff, -dPlate1Thickness - dPlateOff));
                        oPlate13 = new Line3d(new Position(0, 0 - dPlateHorOff + dPlate1Width, -dPlate1Thickness - dPlateOff), new Position(0, 0 - dPlateHorOff, -dPlate1Thickness - dPlateOff));
                        oPlate14 = new Line3d(new Position(0, 0 - dPlateHorOff, -dPlate1Thickness - dPlateOff), new Position(0, 0 - dPlateHorOff, -dPlateOff));


                        oPlate1Coll.Add(oPlate11);
                        oPlate1Coll.Add(oPlate12);
                        oPlate1Coll.Add(oPlate13);
                        oPlate1Coll.Add(oPlate14);

                        ComplexString3d oPlate1ComplxString = new ComplexString3d(oPlate1Coll);
                        Vector Plate1Vector = new Vector(1, 0, 0);
                        Projection3d oPlate1Projection = new Projection3d(oConnection, oPlate1ComplxString, Plate1Vector, dPlate1Depth, true);

                        m_oPhysicalAspect.Outputs["Plate1"] = oPlate1Projection;
                    }
                }

                //Add the ports

                double dHPort1OffsetX = m_dHP1xOffset.Value;
                double dHPort1OffsetY = m_dHP1yOffset.Value;
                double dHPort1OffsetZ = m_dHP1zOffset.Value;
                double dHPort2OffsetX = m_dHP2xOffset.Value;
                double dHPort2OffsetY = m_dHP2yOffset.Value;
                double dHPort2OffsetZ = m_dHP2zOffset.Value;
                double dSportOffsetX = m_dSPxOffset.Value;
                double dSportOffsetY = m_dSPyOffset.Value;
                double dSportOffsetZ = m_dSPzOffset.Value;


                Port oHole1Port = new Port(oConnection, oPart, "Hole1", new Position(dClipPlt1Depth / 2 + dHPort1OffsetX, dClipPlt1Width / 2, dClipPlt1Thickness + dHPort1OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["Hole1"] = oHole1Port;

                //Port oHole2Port = new Port(oConnection, oPart, "Hole2", new Position(dClipPlt1Width / 2 + dHPort2OffsetX, dClipGap / 2 + dClipPlt1Width + dClipPlt2Width + dHorDistance + dHPort2OffsetY, dHPort2OffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                //m_oPhysicalAspect.Outputs["Hole2"] = oHole2Port;
                if (dPlate1Thickness > 0 && dPlate2Thickness > 0)
                {
                    Port oSteelPort = new Port(oConnection, oPart, "Steel", new Position(dClipPlt1Depth / 2 + dSportOffsetX, dClipPlt1Width + dClipPlt2Width + dHorDistance + dSportOffsetY, -dPlate1Thickness - dPlate2Thickness + dSportOffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oPhysicalAspect.Outputs["Steel"] = oSteelPort;
                }
                else if (dPlate1Thickness > 0)
                {
                    Port oSteelPort = new Port(oConnection, oPart, "Steel", new Position(dClipPlt1Depth / 2 + dSportOffsetX, dClipPlt1Width + dClipPlt2Width + dHorDistance + dSportOffsetY, -dPlate1Thickness - dSportOffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oPhysicalAspect.Outputs["Steel"] = oSteelPort;
                }
                else
                {
                    Port oSteelPort = new Port(oConnection, oPart, "Steel", new Position(dClipPlt1Depth / 2 + dSportOffsetX, dClipPlt1Width + dClipPlt2Width + dHorDistance + dSportOffsetY, dSportOffsetZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oPhysicalAspect.Outputs["Steel"] = oSteelPort;
                }



            }
            catch (Exception oExc) //General Unhandled exception 
            {
                throw oExc;
            }
        }
        #endregion

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

