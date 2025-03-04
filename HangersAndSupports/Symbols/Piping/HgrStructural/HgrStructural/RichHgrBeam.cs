//-----------------------------------------------------------------------------
// Copyright 1992 - 2012 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      RichHgrBeam.cs
// Author:     
//      Ramya Pandala     
//
// Abstract:
//     This is .NET RichHgrBeam symbol. This class subclasses from ConnectionComponentDefinition.
//	 22/03/2013		Vijay 	 DI-CP-228142 Modify the error handling for delivered H&S symbols  
//   19/08/2013     Vijay    TR-CP-238035 Weight is wrong when snip functionality is provided for L Rich Hanger Beam.
//   17/04/2014     VDP      CR-CP-245903 Added Reflect Functionality
//   30/12/2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report 
//   19/05/2015      PR       TR 271647  TDL Records created for support objects after data bulkload on Harvest   
//-----------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;
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
    public class RichHgrBeam : ConnectionComponentDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        private const int CUTBACK_STEEL = 1;
        private const int SNIPPED_STEEL = 2;
        private const int NUM_PORTS = 30;
        HangerBeamInputs hgrBeamInput;
        CutbackSteelInputs cutbackSteelInput;
        SnipSteelInputs snipSteelInput;

        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrStructural,Ingr.SP3D.Content.Support.Symbols.RichHgrBeam"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "BeginOverLength", "BeginOverLength of Beam", 0)]
        public InputDouble m_BeginOverLength;
        [InputDouble(3, "EndOverLength", "EndOverLength of the Beam", 0)]
        public InputDouble m_EndOverLength;
        [InputDouble(4, "Length", "Length of the Beam", 0.5)]
        public InputDouble m_Length;
        [InputDouble(5, "CP1", "CP1 of BeginCapPort", 1)]
        public InputDouble m_CP1;
        [InputDouble(6, "CP2", "CP2 of EndCapPort", 1)]
        public InputDouble m_CP2;
        [InputDouble(7, "CP3", "CP3 of BeginFace Port", 5)]
        public InputDouble m_CP3;
        [InputDouble(8, "CP4", "CP4 of EndCapFace Port", 5)]
        public InputDouble m_CP4;
        [InputDouble(9, "CP5", "CP5 of Neutral Port", 5)]
        public InputDouble m_CP5;
        [InputDouble(10, "CP6", "CP6 of Route port", 8)]
        public InputDouble m_CP6;
        [InputDouble(11, "BeginCutbackAnchorPoint", "BeginCutbackAnchorPoint of the Beam", 2)]
        public InputDouble m_BeginCutbackAnchorPoint;
        [InputDouble(12, "EndCutbackAnchorPoint", "EndCutbackAnchorPoint of the Beam", 2)]
        public InputDouble m_EndCutbackAnchorPoint;
        [InputDouble(13, "BeginCapXOffset", "BeginCapXOffset of Beam", 0)]
        public InputDouble m_BeginCapXOffset;
        [InputDouble(14, "BeginCapYOffset", "BeginCapYOffset of the Beam", 0)]
        public InputDouble m_BeginCapYOffset;
        [InputDouble(15, "BeginCapRotZ", "BeginCapRotZ of the Beam", 0)]
        public InputDouble m_BeginCapRotZ;
        [InputDouble(16, "EndCapXOffset", "EndCapXOffset of the Beam", 0)]
        public InputDouble m_EndCapXOffset;
        [InputDouble(17, "EndCapYOffset", "EndCapYOffset of the Beam", 0)]
        public InputDouble m_EndCapYOffset;
        [InputDouble(18, "EndCapRotZ", "EndCapRotZ of the Beam", 1)]
        public InputDouble m_EndCapRotZ;
        [InputDouble(19, "FlexPortXOffset", "FlexPortXOffset of the Beam", 0)]
        public InputDouble m_FlexPortXOffset;
        [InputDouble(20, "FlexPortYOffset", "FlexPortYOffset of the Beam", 0)]
        public InputDouble m_FlexPortYOffset;
        [InputDouble(21, "FlexPortZOffset", "FlexPortZOffset of the Beam", 0)]
        public InputDouble m_FlexPortZOffset;
        [InputDouble(22, "FlexPortRotX", "FlexPortRotX of the Beam", 0)]
        public InputDouble m_FlexPortRotX;
        [InputDouble(23, "FlexPortRotY", "FlexPortRotY of the Beam", 0)]
        public InputDouble m_FlexPortRotY;
        [InputDouble(24, "FlexPortRotZ", "FlexPortRotZ of the Beam", 0)]
        public InputDouble m_FlexPortRotZ;
        [InputDouble(25, "CutbackBeginAngle", "CutbackBeginAngle of the Beam", 0)]
        public InputDouble m_CutbackBeginAngle;
        [InputDouble(26, "CutbackEndAngle", "CutbackEndAngle of the Beam", 0)]
        public InputDouble m_CutbackEndAngle;
        [InputString(27, "MaterialGrade", "MaterialGrade of the Beam", "C20")]
        public InputString m_MaterialGrade;
        [InputString(28, "MaterialType", "MaterialType of the Beam", "Concrete")]
        public InputString m_MaterialType;
        [InputDouble(29, "CoatingType", "CoatingType", 3)]
        public InputDouble m_CoatingType;
        [InputDouble(30, "CoatingRequirement", "CoatingRequirement", 3)]
        public InputDouble m_CoatingRequirement;
        [InputDouble(31, "HgrBeamType", "HgrBeamType", 1)]
        public InputDouble m_HgrBeamType;
        [InputDouble(32, "CutbackBeginAngle1", "CutbackBeginAngle1", 0)]
        public InputDouble m_CutbackBeginAngle1;
        [InputDouble(33, "CutbackBeginAngle2", "CutbackBeginAngle2", 0)]
        public InputDouble m_CutbackBeginAngle2;
        [InputDouble(34, "CutbackEndAngle1", "CutbackEndAngle1", 0)]
        public InputDouble m_CutbackEndAngle1;
        [InputDouble(35, "CutbackEndAngle2", "CutbackEndAngle2", 0)]
        public InputDouble m_CutbackEndAngle2;
        [InputDouble(36, "BeginCutbackAnchorPoint1", "Specifies the edge for the cutback at the begin face of the flange", 0)]
        public InputDouble m_BeginCutbackAnchorPoint1;
        [InputDouble(37, "BeginCutbackAnchorPoint2", "Specifies the edge for the cutback at the begin face of the web", 0)]
        public InputDouble m_BeginCutbackAnchorPoint2;
        [InputDouble(38, "EndCutbackAnchorPoint1", "Specifies the plane for the cutback at the end face of the flange", 0)]
        public InputDouble m_EndCutbackAnchorPoint1;
        [InputDouble(39, "EndCutbackAnchorPoint2", "Specifies the edge for the cutback at the end face of the web", 0)]
        public InputDouble m_EndCutbackAnchorPoint2;
        [InputDouble(40, "BeginOffsetAlongFlange", "Specifies the offset with which the begin face of the flange is snipped", 0)]
        public InputDouble m_BeginOffsetAlongFlange;
        [InputDouble(41, "BeginOffsetAlongWeb", "Specifies the offset with which the begin face of the web is snipped", 0)]
        public InputDouble m_BeginOffsetAlongWeb;
        [InputDouble(42, "EndOffsetAlongFlange", "Specifies the offset with which the end face of the flange is snipped", 0)]
        public InputDouble m_EndOffsetAlongFlange;
        [InputDouble(43, "EndOffsetAlongWeb", "Specifies the offset with which the end face of the web is snipped", 0)]
        public InputDouble m_EndOffsetAlongWeb;
        [InputDouble(44, "FacePortOrient", " Specifies whether the face port orientation is aligned along the flange, or the web when a snip angle is applied", 0)]
        public InputDouble m_FacePortOrient;
        [InputDouble(45, "VarLength", " Specifies the cut length of the steel when you apply cutback angles, snip angles, or overlength on the steel", 0.5)]
        public InputDouble m_VarLength;
        [InputDouble(46, "EndFlexPortXOffset", " Specifies the offset with which the EndFlex port is moved in the x-axis", 0)]
        public InputDouble m_EndFlexPortXOffset;
        [InputDouble(47, "EndFlexPortYOffset", " Specifies the offset with which the EndFlex port is moved in the y-axis", 0)]
        public InputDouble m_EndFlexPortYOffset;
        [InputDouble(48, "EndFlexPortZOffset", " Specifies the offset with which the EndFlex port is moved in the z-axis", 0)]
        public InputDouble m_EndFlexPortZOffset;
        [InputDouble(49, "EndFlexPortRotX", "Specifies the angle with which the EndFlex port is rotated about its x-axis", 0)]
        public InputDouble m_EndFlexPortRotX;
        [InputDouble(50, "EndFlexPortRotY", " Specifies the angle with which the EndFlex port is rotated about its y-axis", 0)]
        public InputDouble m_EndFlexPortRotY;
        [InputDouble(51, "EndFlexPortRotZ", " Specifies the angle with which the EndFlex port is rotated about its z-axis", 0)]
        public InputDouble m_EndFlexPortRotZ;
        [InputDouble(52, "CP7", "CP7 of EndFlex Port", 8)]
        public InputDouble m_CP7;
        [InputDouble(53, "Reflect", "Specifies the reflect plane", 0, true)]
        public InputDouble m_Reflect;
        [InputDouble(54, "ReflectPlaneOffset", "Specifies the offset distance for active reflect plane", 0, true)]
        public InputDouble m_ReflectPlaneOffset;
        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BeginCap", "Begin Cap")]
        [SymbolOutput("EndCap", "End Cap")]
        [SymbolOutput("BeginFace", "Begin Face")]
        [SymbolOutput("EndFace", "End Face")]
        [SymbolOutput("Neutral", "Neutral")]
        [SymbolOutput("BeginFlex", "Begin Flex")]
        [SymbolOutput("BeginCapSurface", "Begin Cap Surface")]
        [SymbolOutput("EndCapSurface", "End Cap Surface")]
        [SymbolOutput("EndFlex", "End Flex")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("Port6", "Port6")]
        [SymbolOutput("Port7", "Port7")]
        [SymbolOutput("Port8", "Port8")]
        [SymbolOutput("Port9", "Port9")]
        [SymbolOutput("Port10", "Port10")]
        [SymbolOutput("Port11", "Port11")]
        [SymbolOutput("Port12", "Port12")]
        [SymbolOutput("Port13", "Port13")]
        [SymbolOutput("Port14", "Port14")]
        [SymbolOutput("Port15", "Port15")]
        [SymbolOutput("Port16", "Port16")]
        [SymbolOutput("Port17", "Port17")]
        [SymbolOutput("Port18", "Port18")]
        [SymbolOutput("Port19", "Port19")]
        [SymbolOutput("Port20", "Port20")]
        [SymbolOutput("Port21", "Port21")]
        [SymbolOutput("Port22", "Port22")]
        [SymbolOutput("Port23", "Port23")]
        [SymbolOutput("Port24", "Port24")]
        [SymbolOutput("Port25", "Port25")]
        [SymbolOutput("Port26", "Port26")]
        [SymbolOutput("Port27", "Port27")]
        [SymbolOutput("Port28", "Port28")]
        [SymbolOutput("Port29", "Port29")]
        [SymbolOutput("Port30", "Port30")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                #region "Definining the variables"

                //Begin Cap Port Orientation
                double xDirX1 = 0, xDirY1 = 0, xDirZ1 = 0, zDirX1 = 0, zDirY1 = 0, zDirZ1 = 0;

                //End Cap Port Orientation
                double xDirX2 = 0, xDirY2 = 0, xDirZ2 = 0, zDirX2 = 0, zDirY2 = 0, zDirZ2 = 0;

                //Begin Face Port Orientation
                double xDirX3 = 0, xDirY3 = 0, xDirZ3 = 0, zDirX3 = 0, zDirY3 = 0, zDirZ3 = 0;

                //End Face Port Orientation
                double xDirX4 = 0, xDirY4 = 0, xDirZ4 = 0, zDirX4 = 0, zDirY4 = 0, zDirZ4 = 0;

                //Route Port Orientation
                double xDirX6 = 0, xDirY6 = 0, xDirZ6 = 0, zDirX6 = 0, zDirY6 = 0, zDirZ6 = 0;

                //Endflex Port Orientation
                double xDirX7 = 0, xDirY7 = 0, xDirZ7 = 0, zDirX7 = 0, zDirY7 = 0, zDirZ7 = 0;

                double snipCBBeginAngle = 0, snipCBEndAngle = 0;
                AnchorPoint snipBeginCBAnchorPt = AnchorPoint.BottomCenter, snipEndCBAnchorPt = AnchorPoint.TopCenter;

                # endregion

                #region "Setting the input values"
                Part sectionPart = (Part)m_PartInput.Value;
                double beginOverLength = 0, endOverLength = 0;
                int cardinalPt1 = 0, cardinalPt2 = 0, cardinalPt3 = 0, cardinalPt4 = 0, cardinalPt5 = 0, cardinalPt6 = 0;
                AnchorPoint beginCutbackAnchorPoint = AnchorPoint.BottomCenter, endCutbackAnchorPoint = AnchorPoint.BottomCenter;
                double beginCapXOffset = 0, beginCapYOffset = 0, beginCapRotZ = 0;
                double endCapXOffset = 0, endCapYOffset = 0, endCapRotZ = 0;
                double flexPortXOffset = 0, flexPortYOffset = 0, flexPortZOffset = 0, flexPortRotX = 0, flexPortRotY = 0, flexPortRotZ = 0;
                double cutbackBeginAngle = 0, cutbackEndAngle = 0;
                double beginAlongFlangeAngle = 0, beginAlongWebAngle = 0, endAlongFlangeAngle = 0, endAlongWebAngle = 0;
                AnchorPoint beginAlongFlangeAnchrPnt = AnchorPoint.BottomCenter, beginAlongWebAnchrPnt = AnchorPoint.BottomCenter, endAlongFlangeAnchrPnt = AnchorPoint.BottomCenter, endAlongWebAnchrPnt = AnchorPoint.BottomCenter;
                double beginAlongFlangeOffset = 0, beginAlongWebOffset = 0, endAlongFlangeOffset = 0, endAlongWebOffset = 0;
                double endFlexPortXOffset = 0, endFlexPortYOffset = 0, endFlexPortZOffset = 0, endFlexPortRotX = 0, endFlexPortRotY = 0, endFlexPortRotZ = 0;
                int cardinalPt7 = 0, bIsCutback = 0;
                int facePortOrient = 0;
                int reflectPlane = 0;
                double reflectPlaneOffset=0;
                string materialType;
                string materialGrade;

                try
                {

                    beginOverLength = m_BeginOverLength.Value;
                    endOverLength = m_EndOverLength.Value;
                    cardinalPt1 = (int)m_CP1.Value;
                    cardinalPt2 = (int)m_CP2.Value;
                    cardinalPt3 = (int)m_CP3.Value;
                    cardinalPt4 = (int)m_CP4.Value;
                    cardinalPt5 = (int)m_CP5.Value;
                    cardinalPt6 = (int)m_CP6.Value;
                    beginCutbackAnchorPoint = (AnchorPoint)m_BeginCutbackAnchorPoint.Value;
                    endCutbackAnchorPoint = (AnchorPoint)m_EndCutbackAnchorPoint.Value;
                    beginCapXOffset = m_BeginCapXOffset.Value;
                    beginCapYOffset = m_BeginCapYOffset.Value;
                    beginCapRotZ = m_BeginCapRotZ.Value;
                    endCapXOffset = m_EndCapXOffset.Value;
                    endCapYOffset = m_EndCapYOffset.Value;
                    endCapRotZ = m_EndCapRotZ.Value;
                    flexPortXOffset = m_FlexPortXOffset.Value;
                    flexPortYOffset = m_FlexPortYOffset.Value;
                    flexPortZOffset = m_FlexPortZOffset.Value;
                    flexPortRotX = m_FlexPortRotX.Value;
                    flexPortRotY = m_FlexPortRotY.Value;
                    flexPortRotZ = m_FlexPortRotZ.Value;
                    cutbackBeginAngle = m_CutbackBeginAngle.Value;
                    cutbackEndAngle = m_CutbackEndAngle.Value;
                    bIsCutback = (int)m_HgrBeamType.Value;
                    beginAlongFlangeAngle = m_CutbackBeginAngle1.Value;
                    beginAlongWebAngle = m_CutbackBeginAngle2.Value;
                    endAlongFlangeAngle = m_CutbackEndAngle1.Value;
                    endAlongWebAngle = m_CutbackEndAngle2.Value;
                    beginAlongFlangeAnchrPnt = (AnchorPoint)m_BeginCutbackAnchorPoint1.Value;
                    beginAlongWebAnchrPnt = (AnchorPoint)m_BeginCutbackAnchorPoint2.Value;
                    endAlongFlangeAnchrPnt = (AnchorPoint)m_EndCutbackAnchorPoint1.Value;
                    endAlongWebAnchrPnt = (AnchorPoint)m_EndCutbackAnchorPoint2.Value;
                    beginAlongFlangeOffset = m_BeginOffsetAlongFlange.Value;
                    beginAlongWebOffset = m_BeginOffsetAlongWeb.Value;
                    endAlongFlangeOffset = m_EndOffsetAlongFlange.Value;
                    endAlongWebOffset = m_EndOffsetAlongWeb.Value;
                    facePortOrient = (int)m_FacePortOrient.Value;
                    endFlexPortXOffset = m_EndFlexPortXOffset.Value;
                    endFlexPortYOffset = m_EndFlexPortYOffset.Value;
                    endFlexPortZOffset = m_EndFlexPortZOffset.Value;
                    endFlexPortRotX = m_EndFlexPortRotX.Value;
                    endFlexPortRotY = m_EndFlexPortRotY.Value;
                    endFlexPortRotZ = m_EndFlexPortRotZ.Value;
                    cardinalPt7 = (int)m_CP7.Value;
                    reflectPlane = (int)m_Reflect.Value;
                    reflectPlaneOffset = m_ReflectPlaneOffset.Value;
                    
                    materialType = m_MaterialType.Value;
                    materialGrade = m_MaterialGrade.Value;
                    if ((cardinalPt1 < 1 || cardinalPt1 > 15) || (cardinalPt2 < 1 || cardinalPt2 > 15) || (cardinalPt3 < 1 || cardinalPt3 > 15) || (cardinalPt4 < 1 || cardinalPt4 > 15) || (cardinalPt5 < 1 || cardinalPt5 > 15) || (cardinalPt6 < 1 || cardinalPt6 > 15) || (cardinalPt7 < 1) || (cardinalPt7 > 15))
                    {
                        if ((cardinalPt1 < 1 || cardinalPt1 > 15))
                            cardinalPt1 = 1;
                        else if (cardinalPt2 < 1 || cardinalPt2 > 15)
                            cardinalPt2 = 1;
                        else if (cardinalPt3 < 1 || cardinalPt3 > 15)
                            cardinalPt3 = 5;
                        else if (cardinalPt4 < 1 || cardinalPt4 > 15)
                            cardinalPt4 = 5;
                        else if (cardinalPt5 < 1 || cardinalPt5 > 15)
                            cardinalPt5 = 5;
                        else if (cardinalPt6 < 1 || cardinalPt6 > 15)
                            cardinalPt6 = 8;
                        else if (cardinalPt7 < 1 || cardinalPt7 > 15)
                            cardinalPt7 = 8;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrInvalidArguments, "Invalid input arguments."));
                    return;
                }

                # endregion

                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                Material material = null;
                double density = 0;

                try
                {
                    material = catalogStructHelper.GetMaterial(materialType, materialGrade);
                    density = material.Density;
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrGettingMaterial, "Error in getting Material Type or Material Grade"));
                    return;
                }

                double cPOffsetx = 0, cPOffsety = 0;

                double length;
                if (m_Length.Value == 0)
                    length = 0.5;
                else
                    length = m_Length.Value;

                #region "Get the Cross Section object from the Relationship"

                //=================================================
                // Construction of Physical Aspect
                //=================================================   
                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection

                CrossSectionServices crossSectionServices = new CrossSectionServices();
                CrossSection crossSection;

                try
                {
                    crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCrossSectionNotFound, "Could not get Cross-section object."));
                    return;
                }

                //Get Section type and standard

                string sectionType = (crossSection.CrossSectionClass.Name);
                double width, depth, flangeThickness, webThickness;

                try
                {
                    GetSectionData(crossSection, out width, out depth, out flangeThickness, out webThickness);
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrSectionNotFound, "Could not get Section data."));
                    return;
                }
                #endregion

                #region "Is Cutback"
                if (bIsCutback == CUTBACK_STEEL)
                {
                    if (cutbackBeginAngle == 0 && cutbackEndAngle == 0)
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
                    else if (beginCutbackAnchorPoint == AnchorPoint.BottomCenter || beginCutbackAnchorPoint == AnchorPoint.TopCenter)
                    {
                        xDirX3 = 1;
                        xDirY3 = 0;
                        xDirZ3 = 0;

                        zDirX3 = 0;
                        zDirY3 = Math.Sin(cutbackBeginAngle);
                        zDirZ3 = Math.Cos(cutbackBeginAngle);

                        if (endCutbackAnchorPoint == AnchorPoint.BottomCenter || endCutbackAnchorPoint == AnchorPoint.TopCenter)
                        {
                            xDirX4 = 1;
                            xDirY4 = 0;
                            xDirZ4 = 0;

                            zDirX4 = 0;
                            zDirY4 = Math.Sin(cutbackEndAngle);
                            zDirZ4 = Math.Cos(cutbackEndAngle);
                        }
                        else
                        {
                            xDirX4 = Math.Cos(cutbackEndAngle);
                            xDirY4 = 0;
                            xDirZ4 = Math.Sin(cutbackEndAngle);

                            zDirX4 = -Math.Sin(cutbackEndAngle);
                            zDirY4 = 0;
                            zDirZ4 = Math.Cos(cutbackEndAngle);
                        }
                    }
                    else if (beginCutbackAnchorPoint == AnchorPoint.MiddleLeft || beginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                    {
                        xDirX3 = Math.Cos(cutbackBeginAngle);
                        xDirY3 = 0;
                        xDirZ3 = Math.Sin(cutbackBeginAngle);

                        zDirX3 = -Math.Sin(cutbackBeginAngle);
                        zDirY3 = 0;
                        zDirZ3 = Math.Cos(cutbackBeginAngle);


                        if (endCutbackAnchorPoint == AnchorPoint.BottomCenter || endCutbackAnchorPoint == AnchorPoint.TopCenter)
                        {
                            xDirX4 = 1;
                            xDirY4 = 0;
                            xDirZ4 = 0;

                            zDirX4 = 0;
                            zDirY4 = Math.Sin(cutbackEndAngle);
                            zDirZ4 = Math.Cos(cutbackEndAngle);
                        }
                        else
                        {

                            xDirX4 = Math.Cos(cutbackEndAngle);
                            xDirY4 = 0;
                            xDirZ4 = Math.Sin(cutbackEndAngle);

                            zDirX4 = -Math.Sin(cutbackEndAngle);
                            zDirY4 = 0;
                            zDirZ4 = Math.Cos(cutbackEndAngle);

                        }
                    }
                }
                #endregion

                #region "Is Snip"
                else
                {
                    #region "Set Angle"
                    if (beginAlongFlangeAngle < 0 || beginAlongWebAngle > 0 || endAlongFlangeAngle > 0 || endAlongWebAngle < 0)
                    {
                        if (beginAlongFlangeAngle < 0)
                            beginAlongFlangeAngle = Math.Abs(beginAlongFlangeAngle);
                        if (beginAlongWebAngle > 0)
                            beginAlongWebAngle = -beginAlongWebAngle;
                        if (endAlongFlangeAngle > 0)
                            endAlongFlangeAngle = -endAlongFlangeAngle;
                        if (endAlongWebAngle < 0)
                            endAlongWebAngle = Math.Abs(endAlongWebAngle);
                    }
                    #endregion

                    #region "Set SnipAngle"
                    if (facePortOrient == 1)
                    {
                        snipCBBeginAngle = -beginAlongFlangeAngle;
                        snipCBEndAngle = -endAlongFlangeAngle;
                        snipBeginCBAnchorPt = beginAlongFlangeAnchrPnt;
                        snipEndCBAnchorPt = endAlongFlangeAnchrPnt;
                    }
                    else if (facePortOrient == 2)
                    {
                        snipCBBeginAngle = beginAlongWebAngle;
                        snipCBEndAngle = endAlongWebAngle;
                        snipBeginCBAnchorPt = beginAlongWebAnchrPnt;
                        snipEndCBAnchorPt = endAlongWebAnchrPnt;
                    }
                    #endregion

                    #region "Set FacePort Orientation for 2"
                    //set the rotation for the face port
                    if (facePortOrient == 2)
                    {
                        if (cardinalPt3 == 1 || cardinalPt3 == 2 || cardinalPt3 == 3 || cardinalPt3 == 11)
                        {
                            if (beginAlongWebOffset > 0)
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = 0;
                                zDirZ3 = 1;
                            }
                            else
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = Math.Sin(snipCBBeginAngle);
                                zDirZ3 = Math.Cos(snipCBBeginAngle);
                            }
                        }
                        else if (cardinalPt3 == 4 || cardinalPt3 == 5 || cardinalPt3 == 6 || cardinalPt3 == 15)
                        {
                            if (beginAlongWebOffset > (depth / 2.0))
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = 0;
                                zDirZ3 = 1;
                            }
                            else
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = Math.Sin(snipCBBeginAngle);
                                zDirZ3 = Math.Cos(snipCBBeginAngle);
                            }
                        }
                        else if (cardinalPt3 == 10 || cardinalPt3 == 12 || cardinalPt3 == 13)
                        {
                            if (beginAlongWebOffset > (depth / 2))
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = 0;
                                zDirZ3 = 1;
                            }
                            else
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = Math.Sin(snipCBBeginAngle);
                                zDirZ3 = Math.Cos(snipCBBeginAngle);
                            }
                        }
                        else if (cardinalPt3 == 7 || cardinalPt3 == 8 || cardinalPt3 == 9 || cardinalPt3 == 14)
                        {
                            xDirX3 = 1;
                            xDirY3 = 0;
                            xDirZ3 = 0;

                            zDirX3 = 0;
                            zDirY3 = Math.Sin(snipCBBeginAngle);
                            zDirZ3 = Math.Cos(snipCBBeginAngle);
                        }

                        if (cardinalPt4 == 1 || cardinalPt4 == 2 || cardinalPt4 == 3 || cardinalPt4 == 11)
                        {
                            if (endAlongWebOffset > 0)
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = 0;
                                zDirZ4 = 1;
                            }
                            else
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = Math.Sin(snipCBEndAngle);
                                zDirZ4 = Math.Cos(snipCBEndAngle);
                            }
                        }
                        else if (cardinalPt4 == 4 || cardinalPt4 == 5 || cardinalPt4 == 6 || cardinalPt4 == 15)
                        {
                            if (endAlongWebOffset > (depth / 2.0))
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = 0;
                                zDirZ4 = 1;
                            }
                            else
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = Math.Sin(snipCBEndAngle);
                                zDirZ4 = Math.Cos(snipCBEndAngle);
                            }

                        }
                        else if (cardinalPt4 == 10 || cardinalPt4 == 12 || cardinalPt4 == 13)
                        {
                            if (endAlongWebOffset > (depth / 2.0))
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = 0;
                                zDirZ4 = 1;
                            }
                            else
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = Math.Sin(snipCBEndAngle);
                                zDirZ4 = Math.Cos(snipCBEndAngle);
                            }
                        }
                        else if (cardinalPt4 == 7 || cardinalPt4 == 8 || cardinalPt4 == 9 || cardinalPt4 == 14)
                        {
                            xDirX4 = 1;
                            xDirY4 = 0;
                            xDirZ4 = 0;

                            zDirX4 = 0;
                            zDirY4 = Math.Sin(snipCBEndAngle);
                            zDirZ4 = Math.Cos(snipCBEndAngle);
                        }
                    }
                    #endregion

                    #region "Set FacePort Orientation for 1"

                    else if (facePortOrient == 1)
                    {
                        if (cardinalPt3 == 1 || cardinalPt3 == 4 || cardinalPt3 == 7 || cardinalPt3 == 12)
                        {
                            if (beginAlongFlangeOffset > 0)
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = 0;
                                zDirZ3 = 1;
                            }
                            else
                            {
                                xDirX3 = Math.Cos(snipCBBeginAngle);
                                xDirY3 = 0;
                                xDirZ3 = -Math.Sin(snipCBBeginAngle);

                                zDirX3 = Math.Sin(snipCBBeginAngle);
                                zDirY3 = 0;
                                zDirZ3 = Math.Cos(snipCBBeginAngle);
                            }
                        }
                        else if (cardinalPt3 == 2 || cardinalPt3 == 5 || cardinalPt3 == 8)
                        {
                            if (beginAlongFlangeOffset > (width / 2.0))
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = 0;
                                zDirZ3 = 1;
                            }
                            else
                            {
                                xDirX3 = Math.Cos(snipCBBeginAngle);
                                xDirY3 = 0;
                                xDirZ3 = -Math.Sin(snipCBBeginAngle);

                                zDirX3 = Math.Sin(snipCBBeginAngle);
                                zDirY3 = 0;
                                zDirZ3 = Math.Cos(snipCBBeginAngle);
                            }
                        }
                        else if (cardinalPt3 == 11 || cardinalPt3 == 10 || cardinalPt3 == 14 || cardinalPt3 == 15)
                        {
                            if (beginAlongFlangeOffset > (width / 4.0))
                            {
                                xDirX3 = 1;
                                xDirY3 = 0;
                                xDirZ3 = 0;

                                zDirX3 = 0;
                                zDirY3 = 0;
                                zDirZ3 = 1;
                            }
                            else
                            {
                                xDirX3 = Math.Cos(snipCBBeginAngle);
                                xDirY3 = 0;
                                xDirZ3 = -Math.Sin(snipCBBeginAngle);

                                zDirX3 = Math.Sin(snipCBBeginAngle);
                                zDirY3 = 0;
                                zDirZ3 = Math.Cos(snipCBBeginAngle);
                            }
                        }
                        else if (cardinalPt3 == 3 || cardinalPt3 == 6 || cardinalPt3 == 9 || cardinalPt3 == 13)
                        {
                            xDirX3 = Math.Cos(snipCBBeginAngle);
                            xDirY3 = 0;
                            xDirZ3 = -Math.Sin(snipCBBeginAngle);

                            zDirX3 = Math.Sin(snipCBBeginAngle);
                            zDirY3 = 0;
                            zDirZ3 = Math.Cos(snipCBBeginAngle);
                        }

                        if (cardinalPt4 == 1 || cardinalPt4 == 4 || cardinalPt4 == 7 || cardinalPt4 == 12)
                        {
                            if (endAlongFlangeOffset > 0)
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = 0;
                                zDirZ4 = 1;
                            }
                            else
                            {
                                xDirX4 = Math.Cos(snipCBEndAngle);
                                xDirY4 = 0;
                                xDirZ4 = -Math.Sin(snipCBEndAngle);

                                zDirX4 = Math.Sin(snipCBEndAngle);
                                zDirY4 = 0;
                                zDirZ4 = Math.Cos(snipCBEndAngle);
                            }
                        }
                        else if (cardinalPt4 == 2 || cardinalPt4 == 5 || cardinalPt4 == 8)
                        {
                            if (endAlongFlangeOffset > (width / 2.0))
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = 0;
                                zDirZ4 = 1;
                            }
                            else
                            {
                                xDirX4 = Math.Cos(snipCBEndAngle);
                                xDirY4 = 0;
                                xDirZ4 = -Math.Sin(snipCBEndAngle);

                                zDirX4 = Math.Sin(snipCBEndAngle);
                                zDirY4 = 0;
                                zDirZ4 = Math.Cos(snipCBEndAngle);
                            }
                        }
                        else if (cardinalPt4 == 11 || cardinalPt4 == 10 || cardinalPt4 == 14 || cardinalPt4 == 15)
                        {
                            if (endAlongFlangeOffset > (width / 4.0))
                            {
                                xDirX4 = 1;
                                xDirY4 = 0;
                                xDirZ4 = 0;

                                zDirX4 = 0;
                                zDirY4 = 0;
                                zDirZ4 = 1;
                            }
                            else
                            {
                                xDirX4 = Math.Cos(snipCBEndAngle);
                                xDirY4 = 0;
                                xDirZ4 = -Math.Sin(snipCBEndAngle);

                                zDirX4 = Math.Sin(snipCBEndAngle);
                                zDirY4 = 0;
                                zDirZ4 = Math.Cos(snipCBEndAngle);
                            }
                        }
                        else if (cardinalPt4 == 3 || cardinalPt4 == 6 || cardinalPt4 == 9 || cardinalPt4 == 13)
                        {
                            xDirX4 = Math.Cos(snipCBEndAngle);
                            xDirY4 = 0;
                            xDirZ4 = -Math.Sin(snipCBEndAngle);

                            zDirX4 = Math.Sin(snipCBEndAngle);
                            zDirY4 = 0;
                            zDirZ4 = Math.Cos(snipCBEndAngle);
                        }
                    }
                    #endregion
                }
                #endregion

                ReadOnlyCollection<BusinessObject> ports = null;
                BusinessObject tempPort;

                #region "Set HgrBeamInput, CutbackSteelInput, SnipSteelInput"

                try
                {
                    hgrBeamInput = new HangerBeamInputs();
                    hgrBeamInput.BeginOverLength = beginOverLength;
                    hgrBeamInput.CardinalPoint = 1;
                    hgrBeamInput.EndOverLength = endOverLength;
                    hgrBeamInput.Length = length;
                    hgrBeamInput.Part = sectionPart;
                    hgrBeamInput.Density = density;
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrInvalidHgrBeamArguments, "Invalid HgrBeamInput data."));
                    return;
                }

                try
                {
                    cutbackSteelInput = new CutbackSteelInputs();
                    cutbackSteelInput.BeginAnchorPoint = beginCutbackAnchorPoint;
                    cutbackSteelInput.BeginOverLength = beginOverLength;
                    cutbackSteelInput.CutbackBeginAngle = cutbackBeginAngle;
                    cutbackSteelInput.CutbackEndAngle = cutbackEndAngle;
                    cutbackSteelInput.Density = density;
                    cutbackSteelInput.EndAnchorPoint = endCutbackAnchorPoint;
                    cutbackSteelInput.EndOverLength = endOverLength;
                    cutbackSteelInput.Length = length;
                    cutbackSteelInput.Part = sectionPart;
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrInvalidCutBackSteelArguments, "Invalid CutbackSteelInput data."));
                    return;
                }

                try
                {
                    snipSteelInput = new SnipSteelInputs();
                    snipSteelInput.BeginFlangeAnchorPoint = beginAlongFlangeAnchrPnt;
                    snipSteelInput.BeginOffsetAlongFlange = beginAlongFlangeOffset;
                    snipSteelInput.BeginOffsetAlongWeb = beginAlongWebOffset;
                    snipSteelInput.BeginOverLength = beginOverLength;
                    snipSteelInput.BeginWebAnchorPoint = beginAlongWebAnchrPnt;
                    snipSteelInput.Density = density;
                    snipSteelInput.EndFlangeAnchorPoint = endAlongFlangeAnchrPnt;
                    snipSteelInput.EndOffsetAlongFlange = endAlongFlangeOffset;
                    snipSteelInput.EndOffsetAlongWeb = endAlongWebOffset;
                    snipSteelInput.EndOverLength = endOverLength;
                    snipSteelInput.EndWebAnchorPoint = endAlongWebAnchrPnt;
                    snipSteelInput.Length = length;
                    snipSteelInput.Part = sectionPart;
                    snipSteelInput.SnipBeginAngleAlongFlange = beginAlongFlangeAngle;
                    snipSteelInput.SnipBeginAngleAlongWeb = beginAlongWebAngle;
                    snipSteelInput.SnipEndAngleAlongFlange = endAlongFlangeAngle;
                    snipSteelInput.SnipEndAngleAlongWeb = endAlongWebAngle;
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrInvalidSnipSteelArguments, "Invalid SnipSteelInput data."));
                    return;
                }
                #endregion

                #region "Check if it is a Cutback Steel or Snip Steel"

                if (bIsCutback == CUTBACK_STEEL)
                {
                    //Cutback not applicable to circular sections
                    if (sectionType == "HSSC" || sectionType == "PIPE" || sectionType == "CS")
                    {
                        try
                        {
                            ports = CreateConnectionComponentPorts(hgrBeamInput);
                        }
                        catch
                        {
                            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateConnectionComponentPortsHgrBeam, "Error in CreateConnectionComponentPorts() for HgrBeamInputs."));
                            return;
                        }
                    }
                    else
                    {
                        try
                        {
                            ports = CreateConnectionComponentPorts(cutbackSteelInput);
                        }
                        catch
                        {
                            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateConnectionComponentPortsCutBackSteel, "Error in CreateConnectionComponentPorts() for CutbackSteelInputs."));
                            return;
                        }
                    }

                }
                else
                {
                    //Snip - only applicable to L section
                    if (sectionType == "L")
                    {
                        try
                        {
                            ports = CreateConnectionComponentPorts(snipSteelInput);
                        }
                        catch
                        {
                            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateConnectionComponentPortsSnipSteel, "Error in CreateConnectionComponentPorts() for SnipSteelInputs."));
                            return;
                        }
                    }
                    else
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrSnipLSectionMaterial, "Snip functionality is supported only for L section"));
                        return;
                    }

                }

                #endregion

                #region "Set BeginCap"
                //Last two ports are reserved for CapSurfaces
                if (ports != null)
                {
                    int portCount = ports.Count - 2;
                    if (portCount > 0)
                    {
                        for (int iIndex = 0; iIndex < portCount; iIndex++)
                        {

                            if (reflectPlane == 1 || reflectPlane == 2)
                            {
                                tempPort = ports[iIndex];
                                Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                                m_PhysicalAspect.Outputs["Port" + (iIndex + 1)] = tempPort;
                            }
                            else
                            {
                                m_PhysicalAspect.Outputs["Port" + (iIndex + 1)] = ports[iIndex]; ;
                            }
                        }
                    }
                    //Now add the Cap Surfaces
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = ports[(int)portCount];
                        Reflect(ref tempPort, crossSection, reflectPlane,reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["BeginCapSurface"] = tempPort;
                        tempPort = ports[(int)portCount + 1];
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["EndCapSurface"] = tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["BeginCapSurface"] = ports[(int)portCount];
                        m_PhysicalAspect.Outputs["EndCapSurface"] = ports[(int)portCount + 1];
                    } 
                }
                //Begin Cap and End Cap ports location offsets
                double x, y, z;

                //Route port location offsets
                double x6, y6, routeZ6;

                //Route port location offsets
                double x7, y7, routeZ7;

                //Begin Face and End Face ports location offsets
                double x1, y1, bFZ1, x2, y2, eFZ2;

                try
                {
                    crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt1, out cPOffsetx, out cPOffsety);//  'Begin Cap

                    x = cPOffsetx + beginCapXOffset;
                    y = cPOffsety + beginCapYOffset;
                    z = 0;

                    if (beginCapRotZ == 0)
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
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrBeginCapPortOrientation, "Error while setting the Orientation for BeginCap Port."));
                    return;
                }

                //BeginCap Port
                try
                {
                    Port port1 = new Port(OccurrenceConnection, sectionPart, "BeginCap", new Position(x, y, z), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = (BusinessObject)port1;
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["BeginCap"] = (Port)tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["BeginCap"] = port1;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateBeginCapPort, "Error while creating BeginCap Port."));
                    return;
                }

                #endregion

                #region "Set EndCap Ports"

                try
                {
                    crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt2, out cPOffsetx, out cPOffsety);//  'End Cap

                    x = cPOffsetx + endCapXOffset;
                    y = cPOffsety + endCapYOffset;
                    z = 0;

                    if (endCapRotZ == 0)
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
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrEndCapPortOrientation, "Error while setting the Orientation for EndCap Port."));
                    return;
                }

                //EndCap Port
                try
                {
                    Port port2 = new Port(OccurrenceConnection, sectionPart, "EndCap", new Position(x, y, z + length), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = (BusinessObject)port2;
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["EndCap"] = (Port)tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["EndCap"] = port2;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateEndCapPort, "Error while creating EndCap Port."));
                    return;
                }
                #endregion

                #region "Set BeginFace"
                try
                {
                    crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt3, out cPOffsetx, out cPOffsety);//    'Begin Face

                    x1 = cPOffsetx;
                    y1 = cPOffsety;

                    if (bIsCutback == CUTBACK_STEEL)
                        bFZ1 = BeginFaceZCoOrdinateForCutback(crossSection, beginCutbackAnchorPoint, cardinalPt3, cutbackBeginAngle);
                    else
                        bFZ1 = BeginFaceZCoOrdinateForSnip(crossSection, cardinalPt3, snipCBBeginAngle, snipBeginCBAnchorPt, facePortOrient, beginAlongFlangeOffset, beginAlongWebOffset);

                    if (sectionType == "HSSC" || sectionType == "PIPE" || sectionType == "CS")
                    {
                        bFZ1 = 0;
                        xDirX3 = 1;
                        xDirY3 = 0;
                        xDirZ3 = 0;
                        zDirX3 = 0;
                        zDirY3 = 0;
                        zDirZ3 = 1;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrBeginFacePortOrientation, "Error while setting the Orientation for BeginFace Port."));
                    return;
                }

                //BeginFace Port
                try
                {
                    Port port3 = new Port(OccurrenceConnection, sectionPart, "BeginFace", new Position(x1, y1, -beginOverLength + bFZ1), new Vector(xDirX3, xDirY3, xDirZ3), new Vector(zDirX3, zDirY3, zDirZ3));
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = (BusinessObject)port3;
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["BeginFace"] = (Port)tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["BeginFace"] = port3;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateBeginFacePort, "Error while creating BeginFace Port."));
                    return;
                }

                #endregion

                #region "Set EndFace"

                try
                {
                    crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt4, out cPOffsetx, out cPOffsety);//    'EndFace

                    x2 = cPOffsetx;
                    y2 = cPOffsety;

                    if (bIsCutback == CUTBACK_STEEL)
                        eFZ2 = EndFaceZCoordForCutback(crossSection, endCutbackAnchorPoint, cardinalPt4, cutbackEndAngle);
                    else
                        eFZ2 = EndFaceZCoordForSnip(crossSection, cardinalPt4, snipCBEndAngle, snipEndCBAnchorPt, facePortOrient, endAlongFlangeOffset, endAlongWebOffset);


                    if (sectionType == "HSSC" || sectionType == "PIPE" || sectionType == "CS")
                    {
                        eFZ2 = 0;
                        xDirX4 = 1;
                        xDirY4 = 0;
                        xDirZ4 = 0;
                        zDirX4 = 0;
                        zDirY4 = 0;
                        zDirZ4 = 1;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrEndFacePortOrientation, "Error while setting the Orientation for EndFace Port."));
                    return;
                }

                //EndFace Port
                try
                {
                    Port port4 = new Port(OccurrenceConnection, sectionPart, "EndFace", new Position(x2, y2, length + endOverLength + eFZ2), new Vector(xDirX4, xDirY4, xDirZ4), new Vector(zDirX4, zDirY4, zDirZ4));
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = (BusinessObject)port4;
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["EndFace"] = (Port)tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["EndFace"] = port4;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateEndFacePort, "Error while creating EndFace Port."));
                    return;
                }

                #endregion

                #region "Neutral Port"
                //Neutral
                try
                {
                    crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt5, out cPOffsetx, out cPOffsety);//    'Neutral

                    x = cPOffsetx;
                    y = cPOffsety;
                    z = length / 2.0;
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrNeutralPortOrientation, "Error while setting the Orientation for Neutral Port."));
                    return;
                }

                try
                {
                    Port port5 = new Port(OccurrenceConnection, sectionPart, "Neutral", new Position(x, y, z), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = (BusinessObject)port5;
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["Neutral"] = (Port)tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["Neutral"] = port5;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateNeutralPort, "Error while creating Neutral Port."));
                    return;
                }

                #endregion

                #region "BeginFlex Port"
                //BeginFlex
                try
                {
                    crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt6, out cPOffsetx, out cPOffsety);

                    x6 = cPOffsetx + flexPortXOffset;
                    y6 = cPOffsety + flexPortYOffset;
                    routeZ6 = flexPortZOffset;

                    if (flexPortRotX == 0 && flexPortRotY == 0 && flexPortRotZ == 0)
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


                        if (flexPortRotX != 0)
                        {
                            tempVactor.Set(1, 0, 0);
                            transMatrix.Rotate(flexPortRotX, tempVactor);
                        }

                        if (flexPortRotY != 0)
                        {
                            tempVactor.Set(0, 1, 0);
                            transMatrix.Rotate(flexPortRotY, tempVactor);
                        }

                        if (flexPortRotZ != 0)
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
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrBeginFlexPortOrientation, "Error while setting the Orientation for BeginFlex Port."));
                    return;
                }

                //BeginFlex
                try
                {
                    Port port6 = new Port(OccurrenceConnection, sectionPart, "BeginFlex", new Position(x6, y6, routeZ6), new Vector(xDirX6, xDirY6, xDirZ6), new Vector(zDirX6, zDirY6, zDirZ6));
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = (BusinessObject)port6;
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["BeginFlex"] = (Port)tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["BeginFlex"] = port6;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateBeginFlexPort, "Error while creating BeginFlex Port."));
                    return;
                }

                #endregion

                #region "EndFlex Port"
                //EndFlex
                try
                {
                    crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt7, out cPOffsetx, out cPOffsety);

                    x7 = cPOffsetx + endFlexPortXOffset;
                    y7 = cPOffsety + endFlexPortYOffset;
                    routeZ7 = length + endFlexPortZOffset;

                    if (endFlexPortRotX == 0 && endFlexPortRotY == 0 && endFlexPortRotZ == 0)
                    {
                        xDirX7 = 1;
                        xDirY7 = 0;
                        xDirZ7 = 0;
                        zDirX7 = 0;
                        zDirY7 = 0;
                        zDirZ7 = 1;
                    }
                    else
                    {
                        Matrix4X4 transMatrix1 = new Matrix4X4();
                        Vector tempVactor1 = new Vector();
                        transMatrix1.SetIdentity();


                        if (endFlexPortRotX != 0)
                        {
                            tempVactor1.Set(1, 0, 0);
                            transMatrix1.Rotate(endFlexPortRotX, tempVactor1);
                        }

                        if (endFlexPortRotY != 0)
                        {
                            tempVactor1.Set(0, 1, 0);
                            transMatrix1.Rotate(endFlexPortRotY, tempVactor1);
                        }

                        if (endFlexPortRotZ != 0)
                        {
                            tempVactor1.Set(0, 0, 1);
                            transMatrix1.Rotate(endFlexPortRotZ, tempVactor1);
                        }
                        xDirX7 = transMatrix1.GetIndexValue(0);
                        xDirY7 = transMatrix1.GetIndexValue(1);
                        xDirZ7 = transMatrix1.GetIndexValue(2);
                        zDirX7 = transMatrix1.GetIndexValue(8);
                        zDirY7 = transMatrix1.GetIndexValue(9);
                        zDirZ7 = transMatrix1.GetIndexValue(10);
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrEndFlexPortOrientation, "Error while setting the Orientation for EndFlex Port."));
                    return;
                }

                //EndFlex
                try
                {
                    Port port7 = new Port(OccurrenceConnection, sectionPart, "EndFlex", new Position(x7, y7, routeZ7), new Vector(xDirX7, xDirY7, xDirZ7), new Vector(zDirX7, zDirY7, zDirZ7));
                    if (reflectPlane == 1 || reflectPlane == 2)
                    {
                        tempPort = (BusinessObject)port7;
                        Reflect(ref tempPort, crossSection, reflectPlane, reflectPlaneOffset);
                        m_PhysicalAspect.Outputs["EndFlex"] = (Port)tempPort;
                    }
                    else
                    {
                        m_PhysicalAspect.Outputs["EndFlex"] = port7;
                    }
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrCreateEndFlexPort, "Error while creating EndFlex Port."));
                    return;
                }
                #endregion

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of RichHgrBeam.cs."));
                    return;
                }
            }

        }
        #endregion

        /// <summary>
        /// Returns the BeginFace Z co-ordinate for cutback steel 
        /// </summary>
        /// <param name="part">RichHgr beam part</param>
        /// <param name="beginCutbackAnchorPoint">Specifies the edge for the cutback at the begin face</param>
        /// <param name="cardinalPoint">Cardinal Point of the cross-section(RichHgrBeam)</param>
        /// <param name="beginCBAngle">Specifies the start cutback angle for the RichHgrBeam</param>
        /// <returns>returns begin face Z coordinate</returns>
        private double BeginFaceZCoOrdinateForCutback(CrossSection crossSection, AnchorPoint beginCutbackAnchorPoint, int cardinalPoint, double beginCBAngle)
        {
            double beginFaceZ = 0;
            AnchorPoint beginCBAnchorPoint = beginCutbackAnchorPoint;
            Part sectionPart = (Part)m_PartInput.Value;
            int beginFaceCardinalPt = cardinalPoint;
            double beginCutbackAngle = beginCBAngle;
            double width, flangeThickness, webThickness, depth;
            double z3, z5;

            GetSectionData(crossSection, out width, out depth, out flangeThickness, out webThickness);

            z3 = width * Math.Tan(-beginCBAngle);    //for Anchor Points 4 and 6
            z5 = depth * Math.Tan(beginCBAngle);    //for Anchor Points 2 and 8

            //Begin Face and End Face Port Offsets
            if (beginCutbackAnchorPoint == AnchorPoint.MiddleLeft || beginCutbackAnchorPoint == AnchorPoint.MiddleRight)
            {
                if (beginFaceCardinalPt == 1 || beginFaceCardinalPt == 4 || beginFaceCardinalPt == 7 || beginFaceCardinalPt == 12)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                        beginFaceZ = 0;
                    else if (beginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                        beginFaceZ = -z3;
                }
                else if (beginFaceCardinalPt == 2 || beginFaceCardinalPt == 5 || beginFaceCardinalPt == 8)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                        beginFaceZ = -(z3 / 2.0);
                    else if (beginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                        beginFaceZ = (z3 / 2.0);

                }
                else if (beginFaceCardinalPt == 11 || beginFaceCardinalPt == 10 || beginFaceCardinalPt == 14 || beginFaceCardinalPt == 15)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                        beginFaceZ = -z3 / 4.0;
                    else if (beginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                        beginFaceZ = z3 / 4.0;
                }
                else if (beginFaceCardinalPt == 3 || beginFaceCardinalPt == 6 || beginFaceCardinalPt == 9 || beginFaceCardinalPt == 13)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.MiddleLeft)
                        beginFaceZ = -z3;
                    else if (beginCutbackAnchorPoint == AnchorPoint.MiddleRight)
                        beginFaceZ = 0;

                }
            }
            else if (beginCutbackAnchorPoint == AnchorPoint.BottomCenter || beginCutbackAnchorPoint == AnchorPoint.TopCenter)
            {
                if (beginFaceCardinalPt == 1 || beginFaceCardinalPt == 2 || beginFaceCardinalPt == 3 || beginFaceCardinalPt == 11)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                        beginFaceZ = 0;
                    else if (beginCutbackAnchorPoint == AnchorPoint.TopCenter)
                        beginFaceZ = z5;
                }
                else if (beginFaceCardinalPt == 4 || beginFaceCardinalPt == 5 || beginFaceCardinalPt == 6 || beginFaceCardinalPt == 15)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                        beginFaceZ = -z5 / 2.0;
                    else if (beginCutbackAnchorPoint == AnchorPoint.TopCenter)
                        beginFaceZ = z5 / 2.0;
                }
                else if (beginFaceCardinalPt == 10 || beginFaceCardinalPt == 12 || beginFaceCardinalPt == 13)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                        beginFaceZ = -z5 / 4.0;
                    else if (beginCutbackAnchorPoint == AnchorPoint.TopCenter)
                        beginFaceZ = z5 / 4.0;
                }
                else if (beginFaceCardinalPt == 7 || beginFaceCardinalPt == 8 || beginFaceCardinalPt == 9 || beginFaceCardinalPt == 14)
                {
                    if (beginCutbackAnchorPoint == AnchorPoint.BottomCenter)
                        beginFaceZ = -z5;
                    else if (beginCutbackAnchorPoint == AnchorPoint.TopCenter)
                        beginFaceZ = 0;

                }
            }
            return beginFaceZ;
        }

        /// <summary>
        /// Returns the BeginFace Z co-ordinate for snip steel 
        /// </summary>
        /// <param name="part">RichHgr beam part</param>
        /// <param name="cardinalPt">Cardinal Point of the cross-section(RichHgrBeam)</param>
        /// <param name="beginCBAngle">Specifies the angle with which the begin face of the RichHgrBeam is snipped</param>
        /// <param name="beginCBAnchorPt">Specifies the edge for the cutback at the begin face of the RichHgrBeam</param>
        /// <param name="facePortOrient">Specifies whether the face port orientation is aligned along the flange, or the web when a snip angle is applied</param>
        /// <param name="beginOffsetAlongFlange">Specifies the offset with which the begin face of the flange is snipped</param>
        /// <param name="beginOffsetAlongWeb">Specifies the offset with which the begin face of the web is snipped</param>
        /// <returns></returns>
        private double BeginFaceZCoOrdinateForSnip(CrossSection crossSection, int cardinalPt, double beginCBAngle, AnchorPoint beginCBAnchorPt, int facePortOrient, double beginOffsetAlongFlange, double beginOffsetAlongWeb)
        {
            double beginFaceZ = 0;
            double width, flangeThickness, webThickness, depth;
            GetSectionData(crossSection, out width, out depth, out flangeThickness, out webThickness);

            AnchorPoint lBeginCBAnchorPt;
            int beginFaceCardinalPt;
            double beginAngleCB, beginOffsetAlongFlange1, beginOffsetAlongWeb1, beginFaceZCoord = 0;

            lBeginCBAnchorPt = beginCBAnchorPt;
            beginAngleCB = beginCBAngle;
            beginFaceCardinalPt = cardinalPt;
            beginOffsetAlongFlange1 = beginOffsetAlongFlange;
            beginOffsetAlongWeb1 = beginOffsetAlongWeb;

            double z3, z5;

            z3 = width * Math.Tan(-beginAngleCB);    //for Anchor Points 4 and 6
            z5 = depth * Math.Tan(beginAngleCB);    //for Anchor Points 2 and 8

            //Begin Face and End Face Port Offsets
            if (facePortOrient == 1)      //for flange cutback
            {
                if (beginFaceCardinalPt == 1 || beginFaceCardinalPt == 4 || beginFaceCardinalPt == 7 || beginFaceCardinalPt == 12)
                {
                    if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                        beginFaceZCoord = 0;
                    else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                        if (beginOffsetAlongFlange1 > 0)
                            beginFaceZCoord = 0;
                        else
                            beginFaceZCoord = z3;
                }
                else if (beginFaceCardinalPt == 2 || beginFaceCardinalPt == 5 || beginFaceCardinalPt == 8)
                {
                    if (beginOffsetAlongFlange1 > (width / 2.0))
                        beginFaceZCoord = 0;
                    else
                        if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                            beginFaceZCoord = -(width / 2.0 - beginOffsetAlongFlange1) * Math.Tan(beginAngleCB);
                        else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                            beginFaceZCoord = width / 2.0 * Math.Tan(-beginAngleCB);
                }
                else if (beginFaceCardinalPt == 11 || beginFaceCardinalPt == 10 || beginFaceCardinalPt == 14 || beginFaceCardinalPt == 15)
                {
                    if (beginOffsetAlongFlange1 > (width / 4.0))
                        beginFaceZCoord = 0;
                    else
                        if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                            beginFaceZCoord = -(width / 4.0 - beginOffsetAlongFlange1) * Math.Tan(-beginAngleCB);
                        else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                            beginFaceZCoord = width / 4.0 * Math.Tan(-beginAngleCB);
                }
                else if (beginFaceCardinalPt == 3 || beginFaceCardinalPt == 6 || beginFaceCardinalPt == 9 || beginFaceCardinalPt == 13)
                {
                    if (lBeginCBAnchorPt == AnchorPoint.MiddleLeft)
                        beginFaceZCoord = (width - beginOffsetAlongFlange1) * Math.Tan(-beginAngleCB);
                    else if (lBeginCBAnchorPt == AnchorPoint.MiddleRight)
                        beginFaceZCoord = 0;
                }
            }
            else if (facePortOrient == 2)      //For web cutback
            {
                if (beginFaceCardinalPt == 1 || beginFaceCardinalPt == 2 || beginFaceCardinalPt == 3 || beginFaceCardinalPt == 11)
                {
                    if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                        beginFaceZCoord = 0;
                    else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)

                        if (beginOffsetAlongWeb1 > 0)
                            beginFaceZCoord = 0;
                        else
                            beginFaceZCoord = z5;
                }
                else if (beginFaceCardinalPt == 4 || beginFaceCardinalPt == 5 || beginFaceCardinalPt == 6 || beginFaceCardinalPt == 15)
                {
                    if (beginOffsetAlongWeb1 > (depth / 2.0))
                        beginFaceZCoord = 0;
                    else
                        if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                            beginFaceZCoord = -(depth / 2.0 - beginOffsetAlongWeb1) * Math.Tan(beginAngleCB);
                        else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)
                            beginFaceZCoord = depth / 2.0 * Math.Tan(beginAngleCB);
                }
                else if (beginFaceCardinalPt == 10 || beginFaceCardinalPt == 12 || beginFaceCardinalPt == 13)
                {
                    if (beginOffsetAlongWeb1 > (depth / 2.0))
                        beginFaceZCoord = 0;
                    else
                        if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                            beginFaceZCoord = -(depth / 4.0 - beginOffsetAlongWeb1) * Math.Tan(beginAngleCB);
                        else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)
                            beginFaceZCoord = depth / 4.0 * Math.Tan(beginAngleCB);
                }
                else if (beginFaceCardinalPt == 7 || beginFaceCardinalPt == 8 || beginFaceCardinalPt == 9 || beginFaceCardinalPt == 14)
                {
                    if (lBeginCBAnchorPt == AnchorPoint.BottomCenter)
                        beginFaceZCoord = -(depth - beginOffsetAlongWeb1) * Math.Tan(beginAngleCB);
                    else if (lBeginCBAnchorPt == AnchorPoint.TopCenter)
                        beginFaceZCoord = 0;
                }
            }
            beginFaceZ = beginFaceZCoord;
            return beginFaceZ;
        }

        /// <summary>
        /// Gets the Cross-Section data
        /// </summary>
        /// <param name="part">RichHgrBeam Part</param>
        /// <param name="width">Width of the cross-section</param>
        /// <param name="depth">Depth of the cross-section</param>
        /// <param name="flangeThickness">Flange thickness of the cross-section</param>
        /// <param name="webThickness">Web thickness of the cross-section</param>
        private void GetSectionData(CrossSection crossSection, out double width, out double depth, out double flangeThickness, out double webThickness)
        {
            flangeThickness = 0;
            webThickness = 0;
            //Get Required Properties From Cross Section
            width = crossSection.Width;
            depth = crossSection.Depth;
            if (crossSection.SupportsInterface("IStructFlangedSectionDimensions"))
            {
                try
                {
                    flangeThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                }
                catch
                {
                    flangeThickness = 0;
                }
                try
                {
                    webThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tw")).PropValue;
                }
                catch
                {
                    webThickness = 0;
                }
            }

            if (flangeThickness <= 0 || webThickness <= 0)
            {
                if (flangeThickness <= 0)
                {
                    flangeThickness = webThickness;
                }
                else
                {
                    webThickness = flangeThickness;
                }
            }
        }

        /// <summary>
        /// Returns the EndFace Z co-ordinate for cutback steel 
        /// </summary>
        /// <param name="part">RichHgr beam part</param>
        /// <param name="endCBAnchorPt">Specifies the edge for the cutback at the end face</param>
        /// <param name="cardinalPt">Cardinal Point of the cross-section(RichHgrBeam)</param>
        /// <param name="endCBAngle">Specifies the end cutback angle for the RichHgrBeam</param>
        /// <returns>returns end face Z coordinate</returns>
        private double EndFaceZCoordForCutback(CrossSection crossSection, AnchorPoint endCBAnchorPt, int cardinalPt, double endCBAngle)
        {
            AnchorPoint lEndCBAnchorPt;
            int endFaceCardinalPt;
            double endFaceZCoord = 0, endAngleCB;

            lEndCBAnchorPt = endCBAnchorPt;
            endAngleCB = endCBAngle;
            endFaceCardinalPt = cardinalPt;

            double width, flangeThickness, webThickness, depth;
            double dZ4, dZ6;

            GetSectionData(crossSection, out width, out depth, out flangeThickness, out webThickness);

            dZ4 = width * Math.Tan(-endAngleCB);   //for Anchor Points 4 and 6
            dZ6 = depth * Math.Tan(endAngleCB);     //for Anchor Points 2 and 8

            if (lEndCBAnchorPt == AnchorPoint.MiddleLeft || lEndCBAnchorPt == AnchorPoint.MiddleRight)
            {
                if (endFaceCardinalPt == 1 || endFaceCardinalPt == 4 || endFaceCardinalPt == 7 || endFaceCardinalPt == 12)
                {
                    if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                        endFaceZCoord = 0;
                    else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                        endFaceZCoord = -dZ4;
                }
                else if (endFaceCardinalPt == 2 || endFaceCardinalPt == 5 || endFaceCardinalPt == 8)
                {
                    if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                        endFaceZCoord = -dZ4 / 2.0;
                    else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                        endFaceZCoord = dZ4 / 2.0;
                }
                else if (endFaceCardinalPt == 11 || endFaceCardinalPt == 10 || endFaceCardinalPt == 14 || endFaceCardinalPt == 15)
                {
                    if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                        endFaceZCoord = -dZ4 / 4.0;
                    else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                        endFaceZCoord = dZ4 / 4.0;
                }
                else if (endFaceCardinalPt == 3 || endFaceCardinalPt == 6 || endFaceCardinalPt == 9 || endFaceCardinalPt == 13)
                {
                    if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                        endFaceZCoord = -dZ4;
                    else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                        endFaceZCoord = 0;
                }
            }
            else if (lEndCBAnchorPt == AnchorPoint.BottomCenter || lEndCBAnchorPt == AnchorPoint.TopCenter)
            {
                if (endFaceCardinalPt == 1 || endFaceCardinalPt == 2 || endFaceCardinalPt == 3 || endFaceCardinalPt == 11)
                {
                    if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                        endFaceZCoord = 0;
                    else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                        endFaceZCoord = dZ6;
                }
                else if (endFaceCardinalPt == 4 || endFaceCardinalPt == 5 || endFaceCardinalPt == 6 || endFaceCardinalPt == 15)
                {
                    if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                        endFaceZCoord = -dZ6 / 2.0;
                    else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                        endFaceZCoord = dZ6 / 2.0;
                }
                else if (endFaceCardinalPt == 10 || endFaceCardinalPt == 12 || endFaceCardinalPt == 13)
                {
                    if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                        endFaceZCoord = -dZ6 / 4.0;
                    else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                        endFaceZCoord = dZ6 / 4.0;
                }
                else if (endFaceCardinalPt == 7 || endFaceCardinalPt == 8 || endFaceCardinalPt == 9 || endFaceCardinalPt == 14)
                {
                    if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                        endFaceZCoord = -dZ6;
                    else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                        endFaceZCoord = 0;
                }
            }
            return endFaceZCoord;
        }

        /// <summary>
        /// Returns the EndFace Z co-ordinate for snip steel 
        /// </summary>
        /// <param name="part">RichHgr beam part</param>
        /// <param name="cardinalPt">Cardinal Point of the cross-section(RichHgrBeam)</param>
        /// <param name="endCBAngle">Specifies the angle with which the end face of the RichHgrBeam is snipped</param>
        /// <param name="endCBAnchorPt">Specifies the edge for the cutback at the end face of the RichHgrBeam</param>
        /// <param name="facePortOrient">Specifies whether the face port orientation is aligned along the flange, or the web when a snip angle is applied</param>
        /// <param name="endOffsetAlongFlange">Specifies the offset with which the end face of the flange is snipped</param>
        /// <param name="EndOffsetAlongWeb">Specifies the offset with which the end face of the web is snipped</param>
        /// <returns></returns>
        private double EndFaceZCoordForSnip(CrossSection crossSection, int cardinalPt, double endCBAngle, AnchorPoint endCBAnchorPt, int facePortOrient, double endOffsetAlongFlange, double EndOffsetAlongWeb)
        {
            AnchorPoint lEndCBAnchorPt;
            int endFaceCardinalPt;
            double endFaceZCoord = 0, endOffsetAlongWeb, endOffsetAlongFlange1, endAngleCB;

            lEndCBAnchorPt = endCBAnchorPt;
            endAngleCB = endCBAngle;
            endFaceCardinalPt = cardinalPt;
            endOffsetAlongWeb = EndOffsetAlongWeb;
            endOffsetAlongFlange1 = endOffsetAlongFlange;

            double width, flangeThickness, webThickness, depth;
            double dZ4, dZ6;

            GetSectionData(crossSection, out width, out depth, out flangeThickness, out webThickness);
            dZ4 = width * Math.Tan(endAngleCB);     //for Anchor Points 4 and 6
            dZ6 = depth * Math.Tan(endAngleCB);     //for Anchor Points 2 and 8

            if (facePortOrient == 1)      //for flange cutback
            {
                if (endFaceCardinalPt == 1 || endFaceCardinalPt == 4 || endFaceCardinalPt == 7 || endFaceCardinalPt == 12)
                {
                    if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                        endFaceZCoord = 0;
                    else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)

                        if (endOffsetAlongFlange1 > 0)
                            endFaceZCoord = 0;
                        else
                            endFaceZCoord = dZ4;
                }
                else if (endFaceCardinalPt == 2 || endFaceCardinalPt == 5 || endFaceCardinalPt == 8)
                {

                    if (endOffsetAlongFlange1 > (width / 2.0))
                        endFaceZCoord = 0;
                    else
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            endFaceZCoord = -(width / 2.0 - endOffsetAlongFlange1) * Math.Tan(endAngleCB);
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            endFaceZCoord = width / 2.0 * Math.Tan(endAngleCB);
                    }
                }
                else if (endFaceCardinalPt == 11 || endFaceCardinalPt == 10 || endFaceCardinalPt == 14 || endFaceCardinalPt == 15)
                {

                    if (endOffsetAlongFlange1 > (width / 4.0))
                        endFaceZCoord = 0;
                    else
                    {
                        if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                            endFaceZCoord = -(width / 4.0 - endOffsetAlongFlange1) * Math.Tan(endAngleCB);
                        else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                            endFaceZCoord = width / 4.0 * Math.Tan(endAngleCB);
                    }
                }
                else if (endFaceCardinalPt == 3 || endFaceCardinalPt == 6 || endFaceCardinalPt == 9 || endFaceCardinalPt == 13)
                {
                    if (lEndCBAnchorPt == AnchorPoint.MiddleLeft)
                        endFaceZCoord = -(width - endOffsetAlongFlange1) * Math.Tan(endAngleCB);
                    else if (lEndCBAnchorPt == AnchorPoint.MiddleRight)
                        endFaceZCoord = 0;
                }
            }
            else if (facePortOrient == 2)      //'for web cutback
            {
                if (endFaceCardinalPt == 1 || endFaceCardinalPt == 2 || endFaceCardinalPt == 3 || endFaceCardinalPt == 11)
                {
                    if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                        endFaceZCoord = 0;
                    else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                    {

                        if (endOffsetAlongWeb > 0)
                            endFaceZCoord = 0;
                        else
                            endFaceZCoord = dZ6;
                    }
                }
                else if (endFaceCardinalPt == 4 || endFaceCardinalPt == 5 || endFaceCardinalPt == 6 || endFaceCardinalPt == 15)
                {
                    if (endOffsetAlongWeb > (depth / 2.0))
                        endFaceZCoord = 0;
                    else
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            endFaceZCoord = -(depth / 2.0 - endOffsetAlongWeb) * Math.Tan(endAngleCB);
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                            endFaceZCoord = depth / 2.0 * Math.Tan(endAngleCB);
                    }
                }
                else if (endFaceCardinalPt == 10 || endFaceCardinalPt == 12 || endFaceCardinalPt == 13)
                {
                    if (endOffsetAlongWeb > (depth / 2.0))
                        endFaceZCoord = 0;
                    else
                    {
                        if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                            endFaceZCoord = -(depth / 4.0 - endOffsetAlongWeb) * Math.Tan(endAngleCB);
                        else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                            endFaceZCoord = depth / 4.0 * Math.Tan(endAngleCB);
                    }
                }
                else if (endFaceCardinalPt == 7 || endFaceCardinalPt == 8 || endFaceCardinalPt == 9 || endFaceCardinalPt == 14)
                {
                    if (lEndCBAnchorPt == AnchorPoint.BottomCenter)
                        endFaceZCoord = -(depth - endOffsetAlongWeb) * Math.Tan(endAngleCB);
                    else if (lEndCBAnchorPt == AnchorPoint.TopCenter)
                        endFaceZCoord = 0;
                }
            }
            return endFaceZCoord;
        }

        /// <summary>
        /// Reflects the Rich Hanger Beam Graphics along the Provided Plane.
        /// </summary>
        /// <param name="businessObject">Business Object of each object</param>
        /// <param name="crossSection">CrossSection of RichHgrBeam</param>
        /// <param name="reflectPlane">Plane of reflection</param>
        /// <param name="reflectPlaneOffset">Reflection plane offset from origin</param>
        /// <returns></returns>
        private void Reflect(ref BusinessObject businessObject,CrossSection crossSection, int reflectPlane,double reflectPlaneOffset)
        {
            ComplexString3d complexString;
            Projection3d projection;
            Point3d point;
            Matrix4X4 translateMatrix = new Matrix4X4();
            Matrix4X4 graphicsReflectMatrix = new Matrix4X4();
            double planeOffsetX;
            double planeOffsetY;
            string sectionType;

            CrossSectionServices crossSectionServices = new CrossSectionServices();
            sectionType = (crossSection.CrossSectionClass.Name);
            crossSectionServices.GetCardinalPointOffset(crossSection, 1, out planeOffsetX, out planeOffsetY);

            //Set Matrix for Translating
            translateMatrix.SetIdentity();
            if (reflectPlane == 1)
            {
                translateMatrix.Translate(new Vector(0, -(2 * planeOffsetY), 0)); //set origin to cardinal point 1 along Y axis
                translateMatrix.Translate(new Vector(0, -(2 * reflectPlaneOffset), 0)); //set plane offset in Y axis
            }
            else
            {
                translateMatrix.Translate(new Vector(-(2 * planeOffsetX), 0, 0)); //set origin to cardinal point 1 along X axis
                translateMatrix.Translate(new Vector(-(2 * reflectPlaneOffset), 0, 0)); //set plane offset in X axis
            }

            //Set Matrix for Reflecting Graphics
            graphicsReflectMatrix.SetIdentity();

            if (reflectPlane == 1)
            {
                graphicsReflectMatrix.SetIndexValue(5, -1);
            }
            else
            {
                graphicsReflectMatrix.SetIndexValue(0, -1);
            }
            graphicsReflectMatrix.MultiplyMatrix(translateMatrix);

            //Reflect graphics based upon type of entity, whether plane or projection
            if (businessObject.SupportsInterface ("IJPlane"))           //Plane
            {
                PlaneHelper planeHelper = new PlaneHelper(businessObject);
                planeHelper.GetBoundary(1, out complexString);
                complexString.Transform(graphicsReflectMatrix);
                planeHelper.SetBoundary(1, complexString);
            }
            else if(businessObject.SupportsInterface("IJProjection"))           //Projection
            {
                if (sectionType == "L")
                {
                    projection = (Projection3d)businessObject;
                    projection.Transform(graphicsReflectMatrix);
                }
                else
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Reflect" + ": " + "ERROR: " + "Circular Sections cannot be reflected", "", "RichHgrBeam.cs", 1978);
                    return;
                }
            }
            else if (businessObject.SupportsInterface("IJPoint"))           //Point
            {
                point = (Point3d)businessObject;
                point.Transform(graphicsReflectMatrix);

            }
            else if (businessObject.SupportsInterface("IJPort"))           //Port
            {
                Port tempPort = (Port)businessObject;
                double[] portOriginArray = new double[4];
                double[] portReflectedValues = new double[4];
                Position portOrigin = new Position();
                int i;
                int j;

                //Get Port Origin and Orientation
                portOrigin = tempPort.Origin;

                portOriginArray[0] = portOrigin.X;
                portOriginArray[1] = portOrigin.Y;
                portOriginArray[2] = portOrigin.Z;
                portOriginArray[3] = 1;

                for (i = 0; i < 4; i++)
                {
                    portReflectedValues[i] = 0;
                    for (j = 0; j < 4; j++)
                    {
                        portReflectedValues[i] = portReflectedValues[i] + (graphicsReflectMatrix.GetIndexValue((j * 4) + i) * portOriginArray[j]);

                    }
                }

                portOrigin.X = portReflectedValues[0];
                portOrigin.Y = portReflectedValues[1];
                portOrigin.Z = portReflectedValues[2];

                //Set Port Origin
                tempPort.Origin = portOrigin;

                businessObject = (BusinessObject)tempPort;

            }
        }

        #region ICustomHgrBOMDescription Members

        /// <summary>
        /// Bill Of Material Description is sethere
        /// </summary>
        /// <param name="supportComponent">SupportComponent for which BOM has to be set</param>
        /// <returns>BOM Description</returns>
        public string BOMDescription(BusinessObject supportComponent)
        {
            string bOMString = "";
            try
            {
                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                CrossSection crossSection;
                Part sectionPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();

                //Get Section type and standard
                string sectionType = crossSection.CrossSectionClass.Name;

                double dblLength = 0, dblCutLen = 0;
                string length = "";

                try
                {
                    dblLength = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAhsVarLength", "VarLength")).PropValue;

                }
                catch
                {
                    dblLength = 0.5;
                    supportComponent.SetPropertyValue(dblLength, "IJUAhsVarLength", "VarLength");
                }

                dblCutLen = dblLength;

                Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)supportComponent.GetRelationship("SupportHasComponents", "Support").TargetObjects[0];
                GenericHelper genericHelper = new GenericHelper(support);
                double unitValue, precision = 0;
                genericHelper.GetDataByRule("HgrStructuralBOMUnits", support, out unitValue);

                if ((UnitName.DISTANCE_METER == (UnitName)unitValue) || ((UnitName)unitValue == UnitName.DISTANCE_MILLIMETER))
                {
                    genericHelper.GetDataByRule("HgrStructuralBOMDecimals", support, out precision);

                }
                if (UnitName.DISTANCE_INCH == (UnitName)unitValue)
                {
                    if (precision > 0)
                        length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dblCutLen, UnitName.DISTANCE_INCH);
                    else
                        length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dblCutLen, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET);
                }
                else
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dblCutLen, (UnitName)unitValue);

                if (sectionType == "HSSC" || sectionType == "PIPE" || sectionType == "CS")
                {
                    bOMString = sectionPart.PartDescription + ", Length: " + length;
                }
                else
                {
                    bOMString = sectionPart.PartDescription + ", Cut Length: " + length;
                }
                return bOMString;

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of RichHgrBeam.cs."));
                }
                return "";
            }
        }

        #endregion

        #region ICustomWeightCG Members

        /// <summary>
        /// Weight and CG is calculated here
        /// </summary>
        /// <param name="supportComponent">support component for which Weight and CG has to be calculated</param>
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                CrossSection crossSection;
                CrossSectionServices crossSectionServices = new CrossSectionServices();
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;

                Part sectionPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double weight = 0, cogX = 0, cogY = 0, cogZ = 0,  reflectPlaneOffset = 0;
                int reflectPlane = 0;
                crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();

                //Get Section type and standard
                string sectionType = (crossSection.CrossSectionClass.Name);
                double beginOverLen = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                double endOverLen = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                double beginCutbackAng = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutback", "CutbackBeginAngle")).PropValue;
                double endCutbackAng = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutback", "CutbackEndAngle")).PropValue;

                double beginAlongFlangeAngle = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackBeginAngle1")).PropValue;
                double beginAlongWebAngle = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackBeginAngle2")).PropValue;
                double endAlongFlangeAngle = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackEndAngle1")).PropValue;
                double endAlongWebAngle = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackEndAngle2")).PropValue;
                double beginAlongFlangeOffset = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "BeginOffsetAlongFlange")).PropValue;
                double beginAlongWebOffset = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "BeginOffsetAlongWeb")).PropValue;
                double endAlongFlangeOffset = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "EndOffsetAlongFlange")).PropValue;
                double endAlongWebOffset = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "EndOffsetAlongWeb")).PropValue;
                double cutLength;
                string materialGrade = (string)((PropertyValueString)supportComponent.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue;
                string materialType = (string)((PropertyValueString)supportComponent.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue;
                if (supportComponent.SupportsInterface("IJOAhsReflect"))
                {
                    reflectPlane = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsReflect", "Reflect")).PropValue;
                    try
                    {
                        reflectPlaneOffset = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsReflect", "ReflectPlaneOffset")).PropValue;
                    }
                    catch
                    {
                        reflectPlaneOffset = 0;
                    }
                }

                double length;

                try
                {
                    length = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                }
                catch
                {
                    length = 0;
                }

                try
                {
                    cutLength = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAhsVarLength", "VarLength")).PropValue;
                }
                catch
                {
                    cutLength = 0;
                }

                int beginAncPoint = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsCutback", "BeginCutbackAnchorPoint")).PropValue;
                int endAncPoint = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsCutback", "EndCutbackAnchorPoint")).PropValue;
                int beginAlongFlangeAnchrPnt = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "BeginCutbackAnchorPoint1")).PropValue;
                int beginAlongWebAnchrPnt = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "BeginCutbackAnchorPoint2")).PropValue;
                int endAlongFlangeAnchrPnt = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "EndCutbackAnchorPoint1")).PropValue;
                int endAlongWebAnchrPnt = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "EndCutbackAnchorPoint2")).PropValue;
                int isCutBack = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAHsHgrBeamType", "HgrBeamType")).PropValue;
                double maxLength = 0;

                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                Material material;
                double density;

                try
                {
                    material = catalogStructHelper.GetMaterial(materialType, materialGrade);

                    density = material.Density;
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrGettingMaterial, "Error in getting Material Type or Material Grade"));
                    return;
                }

                double width, webThk, flangeThk, outerDia, innerDia, area, nominalThk, totLength = 0, cutLen = 0;
                Position centerOfGravity = new Position();

                GetSectionData(crossSection, out width, out outerDia, out flangeThk, out webThk);

                if (length <= 0)
                {
                    length = 0.5;
                    supportComponent.SetPropertyValue(length, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    if (isCutBack == CUTBACK_STEEL)
                    {
                        if (sectionType == "HSSC" || sectionType == "PIPE" || sectionType == "CS")
                        {

                            if (sectionType == "CS")
                            {
                                innerDia = outerDia;
                                area = (Math.PI * 0.25) * (outerDia * outerDia);
                            }
                            else
                            {
                                nominalThk = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IJUAHSS", "tnom")).PropValue;
                                innerDia = outerDia - 2 * nominalThk;
                                area = (Math.PI * 0.25) * (outerDia * outerDia - innerDia * innerDia);
                            }

                            crossSectionServices.GetCardinalPointOffset(crossSection, 10, out cogX, out cogY);

                            totLength = (beginOverLen + length + endOverLen);
                            cogZ = totLength / 2.0;

                            weight = area * totLength * density;
                            cutLen = beginOverLen + length + endOverLen;
                        }
                        else
                        {
                            cutbackSteelInput = new CutbackSteelInputs();
                            cutbackSteelInput.BeginAnchorPoint = (AnchorPoint)beginAncPoint;
                            cutbackSteelInput.BeginOverLength = beginOverLen;
                            cutbackSteelInput.CutbackBeginAngle = beginCutbackAng;
                            cutbackSteelInput.CutbackEndAngle = endCutbackAng;
                            cutbackSteelInput.Density = density;
                            cutbackSteelInput.EndAnchorPoint = (AnchorPoint)endAncPoint;
                            cutbackSteelInput.EndOverLength = endOverLen;
                            cutbackSteelInput.Length = length;
                            cutbackSteelInput.Part = sectionPart;

                            CalculateWeightCG(cutbackSteelInput, ref weight, ref totLength, ref centerOfGravity);
                            cogX = centerOfGravity.X;
                            cogY = centerOfGravity.Y;
                            cogZ = centerOfGravity.Z;

                            if (HgrCompareDoubleService.cmpdbl(length, maxLength) == true)
                            {
                                cutLen = totLength - beginOverLen - endOverLen;
                            }
                            else
                                cutLen = totLength;
                        }

                    }
                    else
                    {
                        if (sectionType == "L")
                        {
                            if (beginAlongFlangeAngle < 0 || beginAlongWebAngle > 0 || endAlongFlangeAngle > 0 || endAlongWebAngle < 0)
                            {
                                if (beginAlongFlangeAngle < 0)
                                {
                                    beginAlongFlangeAngle = Math.Abs(beginAlongFlangeAngle);
                                    supportComponent.SetPropertyValue(beginAlongFlangeAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                                }
                                if (beginAlongWebAngle > 0)
                                {
                                    beginAlongWebAngle = -beginAlongWebAngle;
                                    supportComponent.SetPropertyValue(beginAlongWebAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                                }
                                if (endAlongFlangeAngle > 0)
                                {
                                    endAlongFlangeAngle = -endAlongFlangeAngle;
                                    supportComponent.SetPropertyValue(endAlongFlangeAngle, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                                }
                                if (endAlongWebAngle < 0)
                                {
                                    endAlongWebAngle = Math.Abs(endAlongWebAngle);
                                    supportComponent.SetPropertyValue(endAlongWebAngle, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                                }
                            }

                            cutLen = beginOverLen + length + endOverLen;

                            if (((beginAlongWebOffset > flangeThk && endAlongWebOffset > flangeThk) && (beginAlongFlangeOffset > webThk && endAlongFlangeOffset > webThk)) || ((HgrCompareDoubleService.cmpdbl(beginAlongWebOffset, 0) == false || HgrCompareDoubleService.cmpdbl(endAlongWebOffset, 0) == false || HgrCompareDoubleService.cmpdbl(beginAlongFlangeOffset, 0) == false || HgrCompareDoubleService.cmpdbl(endAlongFlangeOffset, 0) == false) && (HgrCompareDoubleService.cmpdbl(beginAlongFlangeAngle, 0) == true && HgrCompareDoubleService.cmpdbl(beginAlongWebAngle, 0) == true && HgrCompareDoubleService.cmpdbl(endAlongFlangeAngle, 0) == true && HgrCompareDoubleService.cmpdbl(endAlongWebAngle, 0) == true)))
                            {
                                snipSteelInput = new SnipSteelInputs();
                                snipSteelInput.BeginFlangeAnchorPoint = (AnchorPoint)beginAlongFlangeAnchrPnt;
                                snipSteelInput.BeginOffsetAlongFlange = beginAlongFlangeOffset;
                                snipSteelInput.BeginOffsetAlongWeb = beginAlongWebOffset;
                                snipSteelInput.BeginOverLength = beginOverLen;
                                snipSteelInput.BeginWebAnchorPoint = (AnchorPoint)beginAlongWebAnchrPnt;
                                snipSteelInput.Density = density;
                                snipSteelInput.EndFlangeAnchorPoint = (AnchorPoint)endAlongFlangeAnchrPnt;
                                snipSteelInput.EndOffsetAlongFlange = endAlongFlangeOffset;
                                snipSteelInput.EndOffsetAlongWeb = endAlongWebOffset;
                                snipSteelInput.EndOverLength = endOverLen;
                                snipSteelInput.EndWebAnchorPoint = (AnchorPoint)endAlongWebAnchrPnt;
                                snipSteelInput.Length = length;
                                snipSteelInput.Part = sectionPart;
                                snipSteelInput.SnipBeginAngleAlongFlange = beginAlongFlangeAngle;
                                snipSteelInput.SnipBeginAngleAlongWeb = beginAlongWebAngle;
                                snipSteelInput.SnipEndAngleAlongFlange = endAlongFlangeAngle;
                                snipSteelInput.SnipEndAngleAlongWeb = endAlongWebAngle;

                                CalculateWeightCG(snipSteelInput, ref weight, ref cutLen, ref centerOfGravity);

                                cogX = centerOfGravity.X;
                                cogY = centerOfGravity.Y;
                                cogZ = centerOfGravity.Z;
                            }
                            else
                            {
                                CalculateWeightAndCGForSnip(supportComponent, crossSection, out weight, out cogX, out cogY, out cogZ);
                            }
                        }
                    }

                    //Set proper COG values for reflected graphics
                    if(reflectPlane == 1 || reflectPlane==2)
                    {
                        Point3d COG = new Point3d(cogX, cogY, cogZ);
                        BusinessObject tempBO = (BusinessObject)COG;
                        Reflect(ref tempBO, crossSection, reflectPlane, reflectPlaneOffset);
                        COG = (Point3d)tempBO;
                        cogX = COG.X;
                        cogY = COG.Y;
                        cogZ = COG.Z;
                    }

                    supportComponent.SetPropertyValue(cutLen, "IJUAhsVarLength", "VarLength");
                    supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrWeightCG, "Error in WeightCG of RichHgrBeam.cs."));
                }
            }
        }

        /// <summary>
        /// Returns the weight and CG for snip Steel (L Section)
        /// </summary>
        /// <param name="supportComponent">Snip steel for which weight and CG are calculated</param>
        /// <param name="weight">weight of snip Steel(L Section) is returned</param>
        /// <param name="cogX">cogX of snip Steel(L Section) is returned</param>
        /// <param name="cogY">cogY of snip Steel(L Section) is returned</param>
        /// <param name="cogZ">cogZ of snip Steel(L Section) is returned</param>
        void CalculateWeightAndCGForSnip(SupportComponent supportComponent, CrossSection crossSection, out double weight, out double cogX, out double cogY, out double cogZ)
        {
            double webLen, webThk, length, flangeLen, flangeThk;
            double beginOffsetFlng, endOffsetFlng, beginOffsetWeb, endOffsetWeb;
            double beginOffsetFlng_OtherLen, endOffsetFlng_OtherLen, beginOffsetWeb_OtherLen, endOffsetWeb_OtherLen;
            double beginAngleAlongFlng, beginAngleAlongWeb, endAngleAlongFlng, endAngleAlongWeb;
            int facePortOrient;
            double beginLength, endLength, density, totalLen;
            string materialGrade, materialType;

            beginAngleAlongFlng = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackBeginAngle1")).PropValue;
            beginAngleAlongWeb = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackBeginAngle2")).PropValue;
            endAngleAlongFlng = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackEndAngle1")).PropValue;
            endAngleAlongWeb = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsSnipedSteel", "CutbackEndAngle2")).PropValue;

            beginOffsetFlng = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "BeginOffsetAlongFlange")).PropValue;
            endOffsetFlng = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "BeginOffsetAlongWeb")).PropValue;
            beginOffsetWeb = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "EndOffsetAlongFlange")).PropValue;
            endOffsetWeb = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJOAhsCutbackOffset", "EndOffsetAlongWeb")).PropValue;
            try
            {
                length = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
            }
            catch
            {
                length = 0.5;
            }
            facePortOrient = (int)((PropertyValueCodelist)supportComponent.GetPropertyValue("IJOAhsFacePortOrient", "FacePortOrient")).PropValue;

            beginLength = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
            endLength = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
            materialGrade = (string)((PropertyValueString)supportComponent.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue;
            materialType = (string)((PropertyValueString)supportComponent.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue;

            CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
            Material material;

            material = catalogStructHelper.GetMaterial(materialType, materialGrade);
            density = material.Density;

            totalLen = beginLength + length + endLength;

            double cx = 0, cy = 0, cz = 0;

            CrossSectionServices crossSectionServices = new CrossSectionServices();
            crossSectionServices.GetCardinalPointOffset(crossSection, 1, out cx, out cy);

            cz = length / 2.0;
            GetSectionData(crossSection, out flangeLen, out webLen, out flangeThk, out webThk);

            beginOffsetFlng_OtherLen = OtherLength(beginAngleAlongFlng, flangeLen, beginOffsetFlng);
            endOffsetFlng_OtherLen = OtherLength(endAngleAlongFlng, flangeLen, endOffsetFlng);
            beginOffsetWeb_OtherLen = OtherLength(beginAngleAlongWeb, webLen, beginOffsetWeb);
            endOffsetWeb_OtherLen = OtherLength(endAngleAlongWeb, webLen, endOffsetWeb);
            /*     
             ''
             ''    Formula to Calculate Center Of Gravity
             ''
             ''    CgX = [(WL * WT * L * WT/2)] + [(FL-WT) * FT * L * (WT + (FL-WT)/2))] _
             ''           - 1 / 2.0.0 * (FL - BOFL) * S2 * FT * (FL - 1 / 3.0 * (FL - BOFL)) _
             ''           - 1 / 2.0.0 * (WL - BOWL) * S3 * WT * (WT / 2.0.0) _
             ''           - 1 / 2.0.0 * (FL - EOFL) * S4 * FT * (FL - 1 / 3.0 * (FL - EOFL)) _
             ''           - 1 / 2.0.0 * (WL - EOFL) * S5 * WT * (WT / 2.0)   / Total Volume
             ''
             ''
             ''    CgY = [(WL * WT * L * WL/2)] + [(FL-WT) * FT * L * (FT)/2] _
             ''           - 1 / 2.0 * (FL - BOFL) * S2 * FT * (FT / 2.0) _
             ''           - 1 / 2.0 * (WL - BOWL) * S3 * WT * (WL / 2.0) _
             ''           - 1 / 2.0 * (FL - EOFL) * S4 * FT * (FT / 2.0) _
             ''           - 1 / 2.0 * (WL - EOFL) * S5 * WT * (WL / 2.0)  / Total Volume
             ''
             ''    CgZ = [(WL * WT * L * L/2)] + [(FL-WT) * FT * L * L/2] _
             ''           - 1 / 2.0 * (FL - BOFL) * S2 * FT * (FL - 1 / 3.0 * (FL - BOFL)) _
             ''           - 1 / 2.0 * (WL - BOWL) * S3 * WT * (WT / 2.0) _
             ''           - 1 / 2.0 * (FL - EOFL) * S4 * FT * (FL - 1 / 3.0 * (FL - EOFL)) _
             ''           - 1 / 2.0 * (WL - EOFL) * S5 * WT * (WT / 2.0)  / Total Volume
             ''
             ''  Density = UniitWeight / Area
             ''
             ''  TotalWeight = Volume of snip * Density
             ''
             ''
         */
            double webVolume, flangeVolume;
            double beginTriangleAlongFlangeVolume, endTriangleAlongFlangeVolume, beginTriangleAlongWebVolume, endTriangleAlongWebVolume;

            webVolume = webLen * webThk * totalLen;
            flangeVolume = (flangeLen - webThk) * flangeThk * totalLen;

            beginTriangleAlongFlangeVolume = 1.0 / 2.0 * (flangeLen - beginOffsetFlng) * beginOffsetFlng_OtherLen * flangeThk;
            beginTriangleAlongWebVolume = 1.0 / 2.0 * (webLen - beginOffsetWeb) * beginOffsetWeb_OtherLen * webThk;
            endTriangleAlongFlangeVolume = 1.0 / 2.0 * (flangeLen - endOffsetFlng) * endOffsetFlng_OtherLen * flangeThk;
            endTriangleAlongWebVolume = 1.0 / 2.0 * (webLen - endOffsetWeb) * endOffsetWeb_OtherLen * webThk;

            cogX = ((webVolume * webThk / 2.0) + (flangeVolume * (webThk + (flangeLen - webThk) / 2.0)) - beginTriangleAlongFlangeVolume * (flangeLen - 1 / 3.0 * (flangeLen - beginOffsetFlng)) - endTriangleAlongFlangeVolume * (flangeLen - 1 / 3.0 * (flangeLen - endOffsetFlng)) - beginTriangleAlongWebVolume * (webThk / 2.0) - endTriangleAlongWebVolume * webThk / 2.0) / (webVolume + flangeVolume);
            // Transform the values from the origin
            cogX = cogX + cx;
            cogY = ((webVolume * webLen / 2.0) + (flangeVolume * flangeThk / 2.0) - beginTriangleAlongFlangeVolume * flangeThk / 2.0 - endTriangleAlongFlangeVolume * flangeThk / 2.0 - beginTriangleAlongWebVolume * webLen / 2.0 - endTriangleAlongWebVolume * webLen / 2.0) / ((webVolume) + (flangeVolume));
            // Transform the values from the origin
            cogY = cogY + cy;
            cogZ = ((webVolume * totalLen / 2.0) + (flangeVolume * totalLen / 2.0) - beginTriangleAlongFlangeVolume * (flangeLen - 1.0 / 3.0 * (flangeLen - beginOffsetFlng)) - endTriangleAlongFlangeVolume * (flangeLen - 1.0 / 3.0 * (flangeLen - endOffsetFlng)) - beginTriangleAlongWebVolume * (webLen - 1.0 / 3.0 * (webLen - beginOffsetWeb)) - endTriangleAlongWebVolume * (webLen - 1.0 / 3.0 * (webLen - endOffsetWeb))) / (webVolume + flangeVolume);

            // Total Volume of the section
            weight = (webVolume + flangeVolume - beginTriangleAlongFlangeVolume - endTriangleAlongFlangeVolume - beginTriangleAlongWebVolume - endTriangleAlongWebVolume) * density;
        }

        double OtherLength(double angle, double length, double offsetLength)
        {
            try
            {
                return (Math.Abs(Math.Tan(angle) * (length - offsetLength)));
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HangerBeamSymbolLocalizer.GetString(HangerBeamSymbolResourceIDs.ErrWeightCG, "Error in WeightCG of RichHgrBeam.cs."));
                }
                return 0;
            }
        }
        #endregion

    }
}

