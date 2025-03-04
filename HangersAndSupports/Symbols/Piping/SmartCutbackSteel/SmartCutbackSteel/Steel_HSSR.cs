//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Steel_HSSR.cs
//    SmartCutbackSteel,Ingr.SP3D.Content.Support.Symbols.Steel_HSSR
//   Author       :  Vijay
//   Creation Date:  16.11.2012
//   Description: 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16.11.2012     Vijay    CR222284 Converted VB SmartCutBackSteel to C#.Net
//	 22/03/2013		Vijay 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   30.12.2014     PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Linq;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class Steel_HSSR : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "SmartCutbackSteel,Ingr.SP3D.Content.Support.Symbols.Steel_HSSR"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "BeginOverLength", "BeginOverLength", 0.999999)]
        public InputDouble m_BeginOverLength;
        [InputDouble(3, "EndOverLength", "EndOverLength", 0.999999)]
        public InputDouble m_EndOverLength;
        [InputDouble(4, "Length", "Length", 0.999999)]
        public InputDouble m_Length;
        [InputDouble(5, "CP1", "CP1", 0.999999)]
        public InputDouble m_CP1;
        [InputDouble(6, "CP2", "CP2", 0.999999)]
        public InputDouble m_CP2;
        [InputDouble(7, "CP3", "CP3", 0.999999)]
        public InputDouble m_CP3;
        [InputDouble(8, "CP4", "CP4", 0.999999)]
        public InputDouble m_CP4;
        [InputDouble(9, "CP5", "CP5", 0.999999)]
        public InputDouble m_CP5;
        [InputDouble(10, "CP6", "CP6", 0.999999)]
        public InputDouble m_CP6;
        [InputDouble(11, "BeginCutbackAnchorPoint", "BeginCutbackAnchorPoint", 0.999999)]
        public InputDouble m_BeginCutbackAnchorPoint;
        [InputDouble(12, "EndCutbackAnchorPoint", "EndCutbackAnchorPoint", 0.999999)]
        public InputDouble m_EndCutbackAnchorPoint;
        [InputDouble(13, "BeginCapXOffset", "BeginCapXOffset", 0.999999)]
        public InputDouble m_BeginCapXOffset;
        [InputDouble(14, "BeginCapYOffset", "BeginCapYOffset", 0.999999)]
        public InputDouble m_BeginCapYOffset;
        [InputDouble(15, "BeginCapRotZ", "BeginCapRotZ", 0.999999)]
        public InputDouble m_BeginCapRotZ;
        [InputDouble(16, "EndCapXOffset", "EndCapXOffset", 0.999999)]
        public InputDouble m_EndCapXOffset;
        [InputDouble(17, "EndCapYOffset", "EndCapYOffset", 0.999999)]
        public InputDouble m_EndCapYOffset;
        [InputDouble(18, "EndCapRotZ", "EndCapRotZ", 0.999999)]
        public InputDouble m_EndCapRotZ;
        [InputDouble(19, "FlexPortXOffset", "FlexPortXOffset", 0.999999)]
        public InputDouble m_FlexPortXOffset;
        [InputDouble(20, "FlexPortYOffset", "FlexPortYOffset", 0.999999)]
        public InputDouble m_FlexPortYOffset;
        [InputDouble(21, "FlexPortZOffset", "FlexPortZOffset", 0.999999)]
        public InputDouble m_FlexPortZOffset;
        [InputDouble(22, "FlexPortRotX", "FlexPortRotX", 0.999999)]
        public InputDouble m_FlexPortRotX;
        [InputDouble(23, "FlexPortRotY", "FlexPortRotY", 0.999999)]
        public InputDouble m_FlexPortRotY;
        [InputDouble(24, "FlexPortRotZ", "FlexPortRotZ", 0.999999)]
        public InputDouble m_FlexPortRotZ;
        [InputDouble(25, "CutbackBeginAngle", "CutbackBeginAngle", 0.999999)]
        public InputDouble m_CutbackBeginAngle;
        [InputDouble(26, "CutbackEndAngle", "CutbackEndAngle", 0.999999)]
        public InputDouble m_CutbackEndAngle;
        [InputDouble(27, "MaterialGrade", "MaterialGrade", 0.999999)]
        public InputDouble m_oMaterialGrade;
        [InputDouble(28, "MaterialCategory", "MaterialCategory", 0.999999)]
        public InputDouble m_oMaterialCategory;
        [InputDouble(29, "CoatingType", "CoatingType", 0.999999)]
        public InputDouble m_oCoatingType;
        [InputDouble(30, "CoatingRequirement", "CoatingRequirement", 0.999999)]
        public InputDouble m_oCoatingRequirement;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BeginCap", "BeginCap")]
        [SymbolOutput("EndCap", "EndCap")]
        [SymbolOutput("BeginFace", "BeginFace")]
        [SymbolOutput("EndFace", "EndFace")]
        [SymbolOutput("Neutral", "Neutral")]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("BODY1", "BODY1")]
        [SymbolOutput("BODY2", "BODY2")]
        [SymbolOutput("BODY3", "BODY3")]
        [SymbolOutput("BODY4", "BODY4")]
        public AspectDefinition m_Symbolic;


        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
               
                Double beginOverLength = m_BeginOverLength.Value;
                Double endOverLength = m_EndOverLength.Value;
                Double L = m_Length.Value;
                Double beginCapXOffset = m_BeginCapXOffset.Value;
                Double beginCapYOffset = m_BeginCapYOffset.Value;
                Double beginCapRotZ = m_BeginCapRotZ.Value;
                Double endCapXOffset = m_EndCapXOffset.Value;
                Double endCapYOffset = m_EndCapYOffset.Value;
                Double endCapRotZ = m_EndCapRotZ.Value;
                Double flexPortXOffset = m_FlexPortXOffset.Value;
                Double flexPortYOffset = m_FlexPortYOffset.Value;
                Double flexPortZOffset = m_FlexPortZOffset.Value;
                Double flexPortRotX = m_FlexPortRotX.Value;
                Double flexPortRotY = m_FlexPortRotY.Value;
                Double flexPortRotZ = m_FlexPortRotZ.Value;
                Double cutbackBeginAngle = m_CutbackBeginAngle.Value;
                Double cutbackEndAngle = m_CutbackEndAngle.Value;
                long cardinalPt1 = (long)m_CP1.Value;
                long cardinalPt2 = (long)m_CP2.Value;
                long cardinalPt3 = (long)m_CP3.Value;
                long cardinalPt4 = (long)m_CP4.Value;
                long cardinalPt5 = (long)m_CP5.Value;
                long cardinalPt6 = (long)m_CP6.Value;
                long beginCutbackAnchorPoint = (long)m_BeginCutbackAnchorPoint.Value;
                long endCutbackAnchorPoint = (long)m_EndCutbackAnchorPoint.Value;
                if ((cardinalPt1 < 1 || cardinalPt1 > 15) || (cardinalPt2 < 1 || cardinalPt2 > 15) || (cardinalPt3 < 1 || cardinalPt3 > 15) || (cardinalPt4 < 1 || cardinalPt4 > 15) || (cardinalPt5 < 1 || cardinalPt5 > 15) || (cardinalPt6 < 1 || cardinalPt6 > 15))
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrCardinalPnts, "Cardinal points should be between 1 to 15"));
                    return;
                }

                Double Z1 = 0;
                Double Z2 = 0;
                Double Z3 = 0;
                Double Z4 = 0;
                Double Z5 = 0;
                Double Z6 = 0;
                Double Z7 = 0;
                Double Z8 = 0;
                Double cpOffsetX = 0;
                Double cpOffsetY = 0;
                Double X = 0;
                Double Y = 0;
                Double Z = 0;
                Double X6 = 0;
                Double Y6 = 0;
                Double routeZ6 = 0;
                Double X1 = 0;
                Double Y1 = 0;
                Double bfZ1 = 0;
                Double X2 = 0;
                Double Y2 = 0;
                Double efZ2 = 0;

                CrossSection crossSection;
                CrossSectionServices crossSectionServices = new CrossSectionServices();

                try
                {
                    crossSection = (CrossSection)part.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrCrossSectionNotFound, "Unable to get cross section object"));
                    return;
                }

                //Get Section type and standard
                double width, depth, nomThick;
                depth = crossSection.Depth;
                width = crossSection.Width;
                nomThick = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IJUAHSS", "tnom")).PropValue;
                if (HgrCompareDoubleService.cmpdbl(width , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrWidthArgument, "Width cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(nomThick , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrCNomThickArgument, "Wall Thickness cannot be zero"));
                    return;
                }
                //Begin Cap Port Orientation
                Double xDirX1 = 0;
                Double xDirY1 = 0;
                Double xDirZ1 = 0;
                Double zDirX1 = 0;
                Double zDirY1 = 0;
                Double zDirZ1 = 0;

                //End Cap Port Orientation
                Double xDirX2 = 0;
                Double xDirY2 = 0;
                Double xDirZ2 = 0;
                Double zDirX2 = 0;
                Double zDirY2 = 0;
                Double zDirZ2 = 0;

                //Begin Face Port Orientation
                Double xDirX3 = 0;
                Double xDirY3 = 0;
                Double xDirZ3 = 0;
                Double zDirX3 = 0;
                Double zDirY3 = 0;
                Double zDirZ3 = 0;

                //End Face Port Orientation
                Double xDirX4 = 0;
                Double xDirY4 = 0;
                Double xDirZ4 = 0;
                Double zDirX4 = 0;
                Double zDirY4 = 0;
                Double zDirZ4 = 0;

                //Route Port Orientation
                Double xDirX6 = 0;
                Double xDirY6 = 0;
                Double xDirZ6 = 0;
                Double zDirX6 = 0;
                Double zDirY6 = 0;
                Double zDirZ6 = 0;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //for Anchor Points 4 and 6
                Z1 = nomThick * Math.Tan(cutbackBeginAngle);
                Z2 = nomThick * Math.Tan(cutbackEndAngle);
                Z3 = width * Math.Tan(cutbackBeginAngle);
                Z4 = width * Math.Tan(cutbackEndAngle);

                //for Anchor Points 2 and 8
                Z5 = depth * Math.Tan(cutbackBeginAngle);
                Z6 = depth * Math.Tan(cutbackEndAngle);
                Z7 = nomThick * Math.Tan(cutbackBeginAngle);
                Z8 = nomThick * Math.Tan(cutbackEndAngle);

                beginOverLength = beginOverLength * Math.Cos(cutbackBeginAngle);
                endOverLength = endOverLength * Math.Cos(cutbackEndAngle);
                crossSectionServices.GetCardinalPointOffset(crossSection, 1, out cpOffsetX, out cpOffsetY);

                if (HgrCompareDoubleService.cmpdbl(cutbackBeginAngle, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackEndAngle, 0) == true)
                {
                    xDirX3 = 1;
                    xDirY3 = 0;
                    xDirZ3 = 0;
                    zDirX3 = 0;
                    zDirY3 = 0;
                    zDirZ3 = 1;
                    xDirX4 = 1;
                    xDirY4 = 0;
                    xDirZ4 = 0;
                    zDirX4 = 0;
                    zDirY4 = 0;
                    zDirZ4 = 1;
                }
                else if (beginCutbackAnchorPoint == 2 || beginCutbackAnchorPoint == 8 && endCutbackAnchorPoint == 2 || endCutbackAnchorPoint == 8)
                {
                    xDirX3 = 1;
                    xDirY3 = 0;
                    xDirZ3 = 0;
                    zDirX3 = 0;
                    zDirY3 = Math.Sin(cutbackBeginAngle);
                    zDirZ3 = Math.Cos(cutbackBeginAngle);
                    xDirX4 = 1;
                    xDirY4 = 0;
                    xDirZ4 = 0;
                    zDirX4 = 0;
                    zDirY4 = Math.Sin(cutbackEndAngle);
                    zDirZ4 = Math.Cos(cutbackEndAngle);
                }
                else if (beginCutbackAnchorPoint == 4 || beginCutbackAnchorPoint == 6 && endCutbackAnchorPoint == 4 || endCutbackAnchorPoint == 6)
                {
                    xDirX3 = Math.Cos(cutbackBeginAngle);
                    xDirY3 = 0;
                    xDirZ3 = -Math.Sin(cutbackBeginAngle);
                    zDirX3 = Math.Sin(cutbackBeginAngle);
                    zDirY3 = 0;
                    zDirZ3 = Math.Cos(cutbackBeginAngle);
                    xDirX4 = Math.Cos(cutbackEndAngle);
                    xDirY4 = 0;
                    xDirZ4 = -Math.Sin(cutbackEndAngle);
                    zDirX4 = Math.Sin(cutbackEndAngle);
                    zDirY4 = 0;
                    zDirZ4 = Math.Cos(cutbackEndAngle);
                }
                if (beginCutbackAnchorPoint == 4 && endCutbackAnchorPoint == 2 || beginCutbackAnchorPoint == 4 && endCutbackAnchorPoint == 8 || beginCutbackAnchorPoint == 6 && endCutbackAnchorPoint == 2 || beginCutbackAnchorPoint == 6 && endCutbackAnchorPoint == 8 || beginCutbackAnchorPoint == 2 && endCutbackAnchorPoint == 4 || beginCutbackAnchorPoint == 2 && endCutbackAnchorPoint == 6 || beginCutbackAnchorPoint == 8 && endCutbackAnchorPoint == 4 || beginCutbackAnchorPoint == 8 && endCutbackAnchorPoint == 6)
                {
                    beginCutbackAnchorPoint = endCutbackAnchorPoint;
                }
                if (beginCutbackAnchorPoint == 4 && endCutbackAnchorPoint == 4)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, depth - nomThick, -beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, L + endOverLength));
                    pointCollection.Add(new Position(width, depth - nomThick, L + endOverLength - Z4));
                    pointCollection.Add(new Position(width, depth - nomThick, -Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, -beginOverLength));

                    Vector projectionVector = new Vector(0, nomThick, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, -beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength));
                    pointCollection.Add(new Position(nomThick, nomThick, L + endOverLength - nomThick * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(nomThick, nomThick, -nomThick * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, -beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, -beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength));
                    pointCollection.Add(new Position(width, 0, L + endOverLength - Z4));
                    pointCollection.Add(new Position(width, 0, -Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, 0, -beginOverLength));

                    projectionVector = new Vector(0, nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, -(width - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength - (width - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(width, nomThick, L + endOverLength - width * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(width, nomThick, -width * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, -(width - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);
                }
                else if (beginCutbackAnchorPoint == 4 && endCutbackAnchorPoint == 6)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, depth - nomThick, -beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, L + endOverLength + Z4));
                    pointCollection.Add(new Position(width, depth - nomThick, L + endOverLength));
                    pointCollection.Add(new Position(width, depth - nomThick, -Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, -beginOverLength));

                    Vector projectionVector = new Vector(0, nomThick, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, -beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength + Z4));
                    pointCollection.Add(new Position(nomThick, nomThick, L + endOverLength + Z4 - nomThick * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(nomThick, nomThick, -nomThick * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, -beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, -beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength + Z4));
                    pointCollection.Add(new Position(width, 0, L + endOverLength));
                    pointCollection.Add(new Position(width, 0, -Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, 0, -beginOverLength));

                    projectionVector = new Vector(0, nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, -(width - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength + Z4 - (width - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(width, nomThick, L + endOverLength));
                    pointCollection.Add(new Position(width, nomThick, -Z3 - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, -(width - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);
                }
                else if (beginCutbackAnchorPoint == 6 && endCutbackAnchorPoint == 4)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, depth - nomThick, Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, L + endOverLength));
                    pointCollection.Add(new Position(width, depth - nomThick, L + endOverLength - Z4));
                    pointCollection.Add(new Position(width, depth - nomThick, -beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, Z3 - beginOverLength));

                    Vector projectionVector = new Vector(0, nomThick, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength));
                    pointCollection.Add(new Position(nomThick, nomThick, L + endOverLength - nomThick * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(nomThick, nomThick, -nomThick * Math.Tan(cutbackBeginAngle) + Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, Z3 - beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength));
                    pointCollection.Add(new Position(width, 0, L + endOverLength - Z4));
                    pointCollection.Add(new Position(width, 0, -beginOverLength));
                    pointCollection.Add(new Position(0, 0, Z3 - beginOverLength));

                    projectionVector = new Vector(0, nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, nomThick * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength - (width - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(width, nomThick, L + endOverLength - width * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(width, nomThick, -beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, nomThick * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);
                }
                else if (beginCutbackAnchorPoint == 6 && endCutbackAnchorPoint == 6)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, depth - nomThick, Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, L + endOverLength + Z4));
                    pointCollection.Add(new Position(width, depth - nomThick, L + endOverLength));
                    pointCollection.Add(new Position(width, depth - nomThick, -beginOverLength));
                    pointCollection.Add(new Position(0, depth - nomThick, Z3 - beginOverLength));

                    Vector projectionVector = new Vector(0, nomThick, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength + Z4));
                    pointCollection.Add(new Position(nomThick, nomThick, L + endOverLength + Z4 - nomThick * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(nomThick, nomThick, -nomThick * Math.Tan(cutbackBeginAngle) + Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, Z3 - beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, Z3 - beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength + Z4));
                    pointCollection.Add(new Position(width, 0, L + endOverLength));
                    pointCollection.Add(new Position(width, 0, -beginOverLength));
                    pointCollection.Add(new Position(0, 0, Z3 - beginOverLength));

                    projectionVector = new Vector(0, nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, nomThick * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength + Z4 - (width - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(width, nomThick, L + endOverLength));
                    pointCollection.Add(new Position(width, nomThick, -beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, nomThick * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(0, depth - 2 * nomThick, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);
                }
                else if (beginCutbackAnchorPoint == 2 && endCutbackAnchorPoint == 2)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), L + endOverLength - (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, depth, L + endOverLength - depth * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, depth, -depth * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    Vector projectionVector = new Vector(width, 0, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;
                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, -Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength - Z8));
                    pointCollection.Add(new Position(0, (depth - nomThick), L + endOverLength - (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, -Z7 - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, -beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength - Z8));
                    pointCollection.Add(new Position(0, nomThick, -Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, 0, -beginOverLength));

                    projectionVector = new Vector(width, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, -Z7 - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength - Z8));
                    pointCollection.Add(new Position((width - nomThick), (depth - nomThick), L + endOverLength - (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position((width - nomThick), (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, -Z7 - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);
                }
                else if (beginCutbackAnchorPoint == 2 && endCutbackAnchorPoint == 8)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), L + endOverLength + Z8));
                    pointCollection.Add(new Position(0, depth, L + endOverLength));
                    pointCollection.Add(new Position(0, depth, -depth * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    Vector projectionVector = new Vector(width, 0, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, -Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength + (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, (depth - nomThick), L + endOverLength + Z8));
                    pointCollection.Add(new Position(0, (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, -Z7 - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, -beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength + depth * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength + (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, nomThick, -Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, 0, -beginOverLength));

                    projectionVector = new Vector(width, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, -Z7 - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength + (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position((width - nomThick), (depth - nomThick), L + endOverLength + Z8));
                    pointCollection.Add(new Position((width - nomThick), (depth - nomThick), -(depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, -Z7 - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);
                }
                else if (beginCutbackAnchorPoint == 8 && endCutbackAnchorPoint == 2)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, (depth - nomThick), Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), L + endOverLength - (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, depth, L + endOverLength - depth * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, depth, -beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), Z7 - beginOverLength));

                    Vector projectionVector = new Vector(width, 0, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength - Z8));
                    pointCollection.Add(new Position(0, depth - nomThick, L + endOverLength - (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, depth - nomThick, Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, depth * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength - Z8));
                    pointCollection.Add(new Position(0, nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, 0, depth * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(width, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength - Z8));
                    pointCollection.Add(new Position((width - nomThick), depth - nomThick, L + endOverLength - (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position((width - nomThick), depth - nomThick, Z7 - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);
                }
                else if (beginCutbackAnchorPoint == 8 && endCutbackAnchorPoint == 8)
                {
                    Collection<Position> pointCollection = new Collection<Position>();
                    Matrix4X4 transMatrix1 = new Matrix4X4();
                    Vector tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, (depth - nomThick), Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), L + endOverLength + Z8));
                    pointCollection.Add(new Position(0, depth, L + endOverLength));
                    pointCollection.Add(new Position(0, depth, -beginOverLength));
                    pointCollection.Add(new Position(0, (depth - nomThick), Z7 - beginOverLength));

                    Vector projectionVector = new Vector(width, 0, 0);
                    Projection3d projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY1"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength + (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, depth - nomThick, L + endOverLength + Z8));
                    pointCollection.Add(new Position(0, depth - nomThick, Z7 - beginOverLength));
                    pointCollection.Add(new Position(0, nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY2"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position(0, 0, depth * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, 0, L + endOverLength + depth * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, nomThick, L + endOverLength + (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position(0, nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position(0, 0, depth * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(width, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY3"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);

                    pointCollection = new Collection<Position>();
                    transMatrix1 = new Matrix4X4();
                    tempVector1 = new Vector(cpOffsetX, cpOffsetY, 0);
                    transMatrix1.SetIdentity();
                    transMatrix1.Translate(tempVector1);

                    pointCollection.Add(new Position((width - nomThick), nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, L + endOverLength + (depth - nomThick) * Math.Tan(cutbackEndAngle)));
                    pointCollection.Add(new Position((width - nomThick), depth - nomThick, L + endOverLength + Z8));
                    pointCollection.Add(new Position((width - nomThick), depth - nomThick, Z7 - beginOverLength));
                    pointCollection.Add(new Position((width - nomThick), nomThick, (depth - nomThick) * Math.Tan(cutbackBeginAngle) - beginOverLength));

                    projectionVector = new Vector(nomThick, 0, 0);
                    projectionSteel = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY4"] = projectionSteel;

                    projectionSteel.Transform(transMatrix1);
                }
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt1, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX + beginCapXOffset;
                Y = cpOffsetY + beginCapYOffset;
                Z = 0;

                if (HgrCompareDoubleService.cmpdbl(beginCapRotZ , 0)==true)
                {
                    xDirX1 = 1;
                    xDirY1 = 0;
                    xDirZ1 = 0;
                    zDirX1 = 0;
                    zDirY1 = 0;
                    zDirZ1 = 1;
                }
                else
                {
                    xDirX1 = Math.Cos(beginCapRotZ);
                    xDirY1 = Math.Sin(beginCapRotZ);
                    xDirZ1 = 0;
                    zDirX1 = 0;
                    zDirY1 = 0;
                    zDirZ1 = 1;
                }

                Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(X, Y, Z), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                m_Symbolic.Outputs["BeginCap"] = port1;

                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt2, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX + endCapXOffset;
                Y = cpOffsetY + endCapYOffset;
                Z = 0;

                if (HgrCompareDoubleService.cmpdbl(endCapRotZ , 0)==true)
                {
                    xDirX2 = 1;
                    xDirY2 = 0;
                    xDirZ2 = 0;
                    zDirX2 = 0;
                    zDirY2 = 0;
                    zDirZ2 = 1;
                }
                else
                {
                    xDirX2 = Math.Cos(endCapRotZ);
                    xDirY2 = Math.Sin(endCapRotZ);
                    xDirZ2 = 0;
                    zDirX2 = 0;
                    zDirY2 = 0;
                    zDirZ2 = 1;
                }

                Port port2 = new Port(OccurrenceConnection, part, "EndCap", new Position(X, Y, Z + L), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                m_Symbolic.Outputs["EndCap"] = port2;

                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt3, out cpOffsetX, out cpOffsetY);

                X1 = cpOffsetX;
                Y1 = cpOffsetY;
                bfZ1 = GetBeginFaceZCoord(crossSection, beginCutbackAnchorPoint, cardinalPt3, cutbackBeginAngle);

                Port port3 = new Port(OccurrenceConnection, part, "BeginFace", new Position(X1, Y1, -beginOverLength + bfZ1), new Vector(xDirX3, xDirY3, xDirZ3), new Vector(zDirX3, zDirY3, zDirZ3));
                m_Symbolic.Outputs["BeginFace"] = port3;

                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt4, out cpOffsetX, out cpOffsetY);

                X2 = cpOffsetX;
                Y2 = cpOffsetY;
                efZ2 = GetEndFaceZCoord(crossSection, endCutbackAnchorPoint, cardinalPt4, cutbackEndAngle);

                Port port4 = new Port(OccurrenceConnection, part, "EndFace", new Position(X2, Y2, L + endOverLength + efZ2), new Vector(xDirX4, xDirY4, xDirZ4), new Vector(zDirX4, zDirY4, zDirZ4));
                m_Symbolic.Outputs["EndFace"] = port4;

                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt5, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX;
                Y = cpOffsetY;
                Z = L / 2;

                Port port5 = new Port(OccurrenceConnection, part, "Neutral", new Position(X, Y, Z), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Neutral"] = port5;

                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt6, out cpOffsetX, out cpOffsetY);

                X6 = cpOffsetX + flexPortXOffset;
                Y6 = cpOffsetY + flexPortYOffset;
                routeZ6 = L + flexPortZOffset;

                if (HgrCompareDoubleService.cmpdbl(flexPortRotX  , 0)==true && HgrCompareDoubleService.cmpdbl(flexPortRotY  , 0)==true && HgrCompareDoubleService.cmpdbl(flexPortRotZ  , 0)==true)
                {
                    xDirX6 = 1;
                    xDirY6 = 0;
                    xDirZ6 = 0;
                    zDirX6 = 0;
                    zDirY6 = 0;
                    zDirZ6 = 1;
                }
                else
                {
                    Matrix4X4 transMatrix = new Matrix4X4();
                    Vector tempVactor = new Vector();
                    transMatrix.SetIdentity();

                    if (HgrCompareDoubleService.cmpdbl(flexPortRotX  , 0)==false)
                    {
                        tempVactor.Set(1, 0, 0);
                        transMatrix.Rotate(flexPortRotX, tempVactor);
                    }
                    if (HgrCompareDoubleService.cmpdbl(flexPortRotY  , 0)==false)
                    {
                        tempVactor.Set(0, 1, 0);
                        transMatrix.Rotate(flexPortRotY, tempVactor);
                    }
                    if (HgrCompareDoubleService.cmpdbl(flexPortRotZ  , 0)==false)
                    {
                        tempVactor.Set(0, 0, 1);
                        transMatrix.Rotate(flexPortRotZ, tempVactor);
                    }
                    xDirX6 = transMatrix.GetIndexValue(0);
                    xDirY6 = transMatrix.GetIndexValue(1);
                    xDirZ6 = transMatrix.GetIndexValue(2);
                    zDirX6 = transMatrix.GetIndexValue(8);
                    zDirY6 = transMatrix.GetIndexValue(9);
                    zDirZ6 = transMatrix.GetIndexValue(10);
                }

                Port port6 = new Port(OccurrenceConnection, part, "Route", new Position(X6, Y6, routeZ6), new Vector(xDirX6, xDirY6, xDirZ6), new Vector(zDirX6, zDirY6, zDirZ6));
                m_Symbolic.Outputs["Route"] = port6;

                crossSectionServices = null;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrConstructOutputs, "Error in constructoutputs of Steel_HSSR.cs"));
                    return;
                }
            }
        }
        #endregion

        #region "Getting Begin Face Z Coordinate"

        private double GetBeginFaceZCoord(CrossSection crossSection, long beginCutbackAnchorPoint, long cardinalPoint, double beginCutbackAngle)
        {
            double beginFaceZ = 0;
            double Z3, Z5;

            //Get Section type and standard
            double width, depth;
            width = crossSection.Width;
            depth = crossSection.Depth;

            Z3 = width * Math.Tan(beginCutbackAngle);    //for Anchor Points 4 and 6
            Z5 = depth * Math.Tan(beginCutbackAngle);    //for Anchor Points 2 and 8

            if (beginCutbackAnchorPoint == 4 || beginCutbackAnchorPoint == 6)
            {
                if (cardinalPoint == 1 || cardinalPoint == 4 || cardinalPoint == 7 || cardinalPoint == 12)
                {
                    if (beginCutbackAnchorPoint == 4)
                        beginFaceZ = 0;
                    else if (beginCutbackAnchorPoint == 6)
                        beginFaceZ = Z3;
                }
                else if (cardinalPoint == 2 || cardinalPoint == 5 || cardinalPoint == 8 || cardinalPoint == 10 || cardinalPoint == 11 || cardinalPoint == 14 || cardinalPoint == 15)
                {
                    if (beginCutbackAnchorPoint == 4)
                        beginFaceZ = -Z3 / 2;
                    else if (beginCutbackAnchorPoint == 6)
                        beginFaceZ = Z3 / 2;
                }
                else if (cardinalPoint == 3 || cardinalPoint == 6 || cardinalPoint == 9 || cardinalPoint == 13)
                {
                    if (beginCutbackAnchorPoint == 4)
                        beginFaceZ = -Z3;
                    else if (beginCutbackAnchorPoint == 6)
                        beginFaceZ = 0;
                }
            }
            else if (beginCutbackAnchorPoint == 2 || beginCutbackAnchorPoint == 8)
            {
                if (cardinalPoint == 1 || cardinalPoint == 2 || cardinalPoint == 3 || cardinalPoint == 11)
                {
                    if (beginCutbackAnchorPoint == 2)
                        beginFaceZ = 0;
                    else if (beginCutbackAnchorPoint == 8)
                        beginFaceZ = Z5;
                }
                else if (cardinalPoint == 4 || cardinalPoint == 5 || cardinalPoint == 6 || cardinalPoint == 10 || cardinalPoint == 12 || cardinalPoint == 13 || cardinalPoint == 15)
                {
                    if (beginCutbackAnchorPoint == 2)
                        beginFaceZ = -Z5 / 2;
                    else if (beginCutbackAnchorPoint == 8)
                        beginFaceZ = Z5 / 2;
                }
                else if (cardinalPoint == 7 || cardinalPoint == 8 || cardinalPoint == 9 || cardinalPoint == 14)
                {
                    if (beginCutbackAnchorPoint == 2)
                        beginFaceZ = -Z5;
                    else if (beginCutbackAnchorPoint == 8)
                        beginFaceZ = 0;
                }
            }
            return beginFaceZ;
        }
        #endregion

        #region "Getting End Face Z Coordinate"

        private double GetEndFaceZCoord(CrossSection crossSection, long endCutbackAnchorPoint, long cardinalPoint, double endCutbackAngle)
        {
            double endFaceZ = 0;
            double Z4, Z6;

            //Get Section type and standard
            double width, depth;
            width = crossSection.Width;
            depth = crossSection.Depth;

            Z4 = width * Math.Tan(endCutbackAngle);    //for Anchor Points 4 and 6
            Z6 = depth * Math.Tan(endCutbackAngle);

            if (endCutbackAnchorPoint == 4 || endCutbackAnchorPoint == 6)
            {
                if (cardinalPoint == 1 || cardinalPoint == 4 || cardinalPoint == 7 || cardinalPoint == 12)
                {
                    if (endCutbackAnchorPoint == 4)
                        endFaceZ = 0;
                    else if (endCutbackAnchorPoint == 6)
                        endFaceZ = Z4;
                }
                else if (cardinalPoint == 2 || cardinalPoint == 5 || cardinalPoint == 8 || cardinalPoint == 10 || cardinalPoint == 11 || cardinalPoint == 14 || cardinalPoint == 15)
                {
                    if (endCutbackAnchorPoint == 4)
                        endFaceZ = -Z4 / 2;
                    else if (endCutbackAnchorPoint == 6)
                        endFaceZ = Z4 / 2;
                }
                else if (cardinalPoint == 3 || cardinalPoint == 6 || cardinalPoint == 9 || cardinalPoint == 13)
                {
                    if (endCutbackAnchorPoint == 4)
                        endFaceZ = -Z4;
                    else if (endCutbackAnchorPoint == 6)
                        endFaceZ = 0;
                }
            }
            else if (endCutbackAnchorPoint == 2 || endCutbackAnchorPoint == 8)
            {
                if (cardinalPoint == 1 || cardinalPoint == 2 || cardinalPoint == 3 || cardinalPoint == 11)
                {
                    if (endCutbackAnchorPoint == 2)
                        endFaceZ = 0;
                    else if (endCutbackAnchorPoint == 8)
                        endFaceZ = Z6;
                }
                else if (cardinalPoint == 4 || cardinalPoint == 5 || cardinalPoint == 6 || cardinalPoint == 10 || cardinalPoint == 12 || cardinalPoint == 13 || cardinalPoint == 15)
                {
                    if (endCutbackAnchorPoint == 2)
                        endFaceZ = -Z6 / 2;
                    else if (endCutbackAnchorPoint == 8)
                        endFaceZ = Z6 / 2;
                }
                else if (cardinalPoint == 7 || cardinalPoint == 8 || cardinalPoint == 9 || cardinalPoint == 14)
                {
                    if (endCutbackAnchorPoint == 2)
                        endFaceZ = -Z6;
                    else if (endCutbackAnchorPoint == 8)
                        endFaceZ = 0;
                }
            }
            return endFaceZ;
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                CommonSmartCutBackSteel commonSmartCutBackSteel = new CommonSmartCutBackSteel();
                PropertyValueCodelist beginAncPointValue, endAncPointValue;
                CodelistItem codeBeginAncpoint, codeEndAncPoint;
                double beginOverLenValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                double endOverLenValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                beginAncPointValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAhsCutback", "BeginCutbackAnchorPoint");
                codeBeginAncpoint = beginAncPointValue.PropertyInfo.CodeListInfo.GetCodelistItem(beginAncPointValue.PropValue);
                endAncPointValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAhsCutback", "EndCutbackAnchorPoint");
                codeEndAncPoint = endAncPointValue.PropertyInfo.CodeListInfo.GetCodelistItem(endAncPointValue.PropValue);
                double beginCutbackAngValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAhsCutback", "CutbackBeginAngle")).PropValue;
                double endCutbackAngValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAhsCutback", "CutbackEndAngle")).PropValue;
                double LValue = 0;
                double length = 0;

                length = commonSmartCutBackSteel.CalculateLengthForBOM(part, LValue, beginOverLenValue, endOverLenValue, codeBeginAncpoint, codeEndAncPoint, beginCutbackAngValue, endCutbackAngValue);
                Ingr.SP3D.Support.Middle.Support oSupport = (Ingr.SP3D.Support.Middle.Support)oSupportOrComponent.GetRelationship("SupportHasComponents", "Support").TargetObjects[0];

                GenericHelper genericHelper = new GenericHelper(oSupport);
                double unitValue;
                genericHelper.GetDataByRule("HgrStructuralBOMUnits", oSupport, out unitValue);
                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_INCH);
                bomString = part.PartDescription + ", Overall Length: " + L;
                return bomString;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Steel_HSSR.cs"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Double weight, cogX, cogY, cogZ;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                CommonSmartCutBackSteel commonSmartCutBackSteel = new CommonSmartCutBackSteel();
                PropertyValueCodelist beginAncPointValue, endAncPointValue;
                CodelistItem codeBeginAncpoint, codeEndAncPoint;
                double beginOverLenValue = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                double endOverLenValue = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                beginAncPointValue = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAhsCutback", "BeginCutbackAnchorPoint");
                codeBeginAncpoint = beginAncPointValue.PropertyInfo.CodeListInfo.GetCodelistItem(beginAncPointValue.PropValue);
                endAncPointValue = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAhsCutback", "EndCutbackAnchorPoint");
                codeEndAncPoint = endAncPointValue.PropertyInfo.CodeListInfo.GetCodelistItem(endAncPointValue.PropValue);
                double beginCutbackAngValue = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAhsCutback", "CutbackBeginAngle")).PropValue;
                double endCutbackAngValue = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAhsCutback", "CutbackEndAngle")).PropValue;
                double LValue = 0, beginFaceZ, endFaceZ;
                double density = 0.25;
                beginOverLenValue = beginOverLenValue * Math.Cos(beginCutbackAngValue);
                endOverLenValue = endOverLenValue * Math.Cos(endCutbackAngValue);

                CrossSection crossSection;
                CrossSectionServices crossSectionServices = new CrossSectionServices();

                try
                {
                    crossSection = (CrossSection)part.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrCrossSectionNotFound, "Unable to get cross section object"));
                    return;
                }
                crossSectionServices.GetCardinalPointOffset(crossSection, 10, out cogX, out cogY);
                beginFaceZ = GetBeginFaceZCoord(crossSection, codeBeginAncpoint.Value, 5, beginCutbackAngValue);
                endFaceZ = GetEndFaceZCoord(crossSection, codeEndAncPoint.Value, 5, endCutbackAngValue);
                double length = 0;
                length = commonSmartCutBackSteel.CalculateLengthForWCG(crossSection, LValue, beginOverLenValue, endOverLenValue, codeBeginAncpoint, codeEndAncPoint, beginCutbackAngValue, endCutbackAngValue, beginFaceZ, endFaceZ);
                cogZ = length / 2.0;
                weight = density * length;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartCutbackSteelLocalizer.GetString(SmartCutbackSteelSymbolResourceIDs.ErrWeightCG, "Error in weightCG of Steel_HSSR.cs"));
                }
            }
        }
        #endregion
    }
}

