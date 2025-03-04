
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//
//Copyright 1992 - 2014 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  HandRailTypeA.vb
//
//Abstract
//	This is .NET HandRailTypeA symbol. This class subclasses from HandRailSymbolDefinition.
//
 
//History:
//      July 09, 2012   3XCalibur           CR184504 GetCrossSectionDimensions should be moved from SymbolHelper to CrossSectionService
//                                          Handled the impact for Width and Depth properties on the CrossSection
//      Feb 16, 2015    3XCalibur           DI-CP-267808  Implement content changes for support of drop of Handrail 
//
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Exceptions;
using System;
using System.Collections.Generic;

//===========================================================================================
//Namespace of this class is Ingr.SP3D.Content.Structure
//It is recommended that customers specify namespace of their symbols to be
//<CompanyName>.SP3D.Content.<Specialization>.
//It is also recommended that if customers want to change this symbol to suit their
//requirements, they should change namespace/symbol name so the identity of the modified
//symbol will be different from the one delivered by Intergraph.
//===========================================================================================

namespace Ingr.SP3D.Content.Structure
{
    ///// <summary>
    ///// 
    ///// </summary>
    //internal class HandRailPropertyValues
    //{

    //}
    /// <summary>
    /// // DefinitionName/ProgID of this symbol is "HandRailSymbols,Ingr.SP3D.Content.Structure.HandRailTypeA"
    /// </summary>
    public class HandRailTypeA : HandRailSymbolDefinition, ICustomWeightCG, ICustomConversionToComponents
    {
        #region Definition of Inputs
        private const string HRBeginTreatment = "HRBeginTreatment";
        private const string HREndTreatment = "HREndTreatment";
        private HandRailPropertyValues handrailPropertyValues;
        private HandRail handRail;
        private bool IsErrorToDoMessageCreated
        {
            get
            {
                // Check if to-do message is created while validating handrail property values.
                if (base.ToDoListMessage != null)
                {
                    // Return true if it is an error message.
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        [InputCatalogPart(1)]
        public InputCatalogPart catalogPart;
        [InputObject(2, "Sketch3d", "Sketch3d output complex curve")]
        public InputObject sketch3d;
        [InputDouble(3, "Height", "Height of the Handrail from path to top of Top Rail", 1.067)]
        public InputDouble height;
        [InputDouble(4, "WithToePlate", "With or Without Toe/Kick plate", 1.0)]
        public InputDouble withToePlate;
        [InputDouble(5, "NoOfMidRails", "With One/Many/None mid rails", 1.0)]
        public InputDouble numberOfMidrails;
        [InputDouble(6, "HorizontalOffset", "Offset position that the handrail will be from the path chosen", 15.0)]
        public InputDouble horizontalOffsetPosition;
        [InputDouble(7, "HorizontalOffsetDim", "Distance that the handrail will be offset from the path chosen", 0.0762)]
        public InputDouble horizontalOffsetDimension;
        [InputDouble(8, "HandrailOrientation", "Vertical / Normal", 0.0)]
        public InputDouble handrailOrientation;
        [InputDouble(9, "SegmentMaxSpacing", "Maximum spacing between posts on straight/arc line", 1.524)]
        public InputDouble segmentMaximumSpacing;
        [InputDouble(10, "SlopedSegmentMaxSpacing", "Maximum spacing between posts on slope", 1.829)]
        public InputDouble slopedSegmentMaximumSpacing;
        [InputDouble(11, "TopOfToePlateDim", "Dimension to the top of the toe plate from Path", 0.102)]
        public InputDouble topOfToePlateDimension;
        [InputDouble(12, "TopOfMidRailDim", "Dimension to the top of the Mid rail from top of toe plate", 0.5334)]
        public InputDouble topOfMidrailDimension;
        [InputDouble(13, "MidRailSpacing", "Spacing between mid rails", 0.2)]
        public InputDouble midrailSpacing;
        [InputDouble(14, "WithPostAtTurn", "With or Without Post at the turn", 1.0)]
        public InputDouble withPostAtTurn;
        [InputDouble(15, "BeginTreatmentType", "Begin treatment type 0 for rectangular, 1 for rounded", 2.0)]
        public InputDouble beginTreatmentType;
        [InputDouble(16, "BeginExtensionLength", "Extent of the end of the rail (at beginning)", 0.6096)]
        public InputDouble beginExtensionLength;
        [InputDouble(17, "EndTreatmentType", "End treatment type 0 for rectangular, 1 for rounded", 2.0)]
        public InputDouble endTreatmentType;
        [InputDouble(18, "EndExtensionLength", "Extent of the end of the rail (at ending)", 0.6096)]
        public InputDouble endExtensionLength;
        [InputDouble(19, "PostConnectionType", "Connection required to assemble the entire handrail", 1.0)]
        public InputDouble postConnectionType;
        [InputDouble(20, "IsAssembly", "If system placed as a symbol otherwise constructed with members", 0.0)]
        public InputDouble isAssembly;
        [InputDouble(21, "TopRailSectionCP", "TopRail Section Cardinal Point", 5.0)]
        public InputDouble topRailSectionCardinalPoint;
        [InputDouble(22, "TopRailSectionAngle", "TopRail Section Angle", 0)]
        public InputDouble toprailSectionAngle;
        [InputDouble(23, "MidRailSectionCP", "MidRail Section Cardinal Point", 5.0)]
        public InputDouble midrailSectionCardinalPoint;
        [InputDouble(24, "MidRailSectionAngle", "MidRail Section Angle", 0.0)]
        public InputDouble midrailSectionAngle;
        [InputDouble(25, "ToePlateSectionCP", "ToePlate Section Cardinal Point", 2.0)]
        public InputDouble toePlateSectionCardinalPoint;
        [InputDouble(26, "ToePlateSectionAngle", "ToePlate Section Angle", 0.0)]
        public InputDouble toePlateSectionAngle;
        [InputDouble(27, "PostSectionCP", "Post Section Cardinal Point", 5.0)]
        public InputDouble postSectionCardinalPoint;
        [InputDouble(28, "PostSectionAngle", "Post Section Angle", 0.0)]
        public InputDouble postSectionAngle;
        [InputString(29, "TopRail_SPSSectionName", "TopRail Section from structural cross sections", "HSS1-7/8x.145")]
        public InputString toprailSectionName;
        [InputString(30, "TopRail_SPSSectionRefStandard", "TopRail Section Reference Standard", "AISC-LRFD-3.1")]
        public InputString toprailSectionReferenceStandard;
        [InputString(31, "MidRail_SPSSectionName", "MidRail Section from structural cross sections", "HSS1-7/8x.145")]
        public InputString midrailSectionName;
        [InputString(32, "MidRail_SPSSectionRefStandard", "MidRail Section Reference Standard", "AISC-LRFD-3.1")]
        public InputString midrailSectionReferenceStandard;
        [InputString(33, "ToePlate_SPSSectionName", "ToePlate Section from structural cross sections", "RS0.125x4")]
        public InputString toePlateSectionName;
        [InputString(34, "ToePlate_SPSSectionRefStandard", "ToePlate Section Reference Standard", "Misc")]
        public InputString toePlateSectionReferenceStandard;
        [InputString(35, "Post_SPSSectionName", "Post Section from structural cross sections", "HSS1-7/8x.145")]
        public InputString postSectionName;
        [InputString(36, "Post_SPSSectionRefStandard", "Post Section Reference Standard", "AISC-LRFD-3.1")]
        public InputString postSectionReferenceStandard;
        [InputString(37, "Primary_SPSMaterial", "Primary Material", "Steel - Carbon")]
        public InputString primaryMaterial;
        [InputString(38, "Primary_SPSGrade", "Primary Material Grade", "A")]
        public InputString primaryMaterialGrade;

        #endregion

        #region Definitions of Aspects and their outputs

        //SimplePhysical Aspect
        [Aspect("SimplePhysical", "SimplePhysical Aspect of Handrail", AspectID.SimplePhysical)]

        public AspectDefinition simplePhysicalAspect;
        //Operation Aspect
        [Aspect("Operation", "Operation Aspect of Handrail", AspectID.Operation)]
        [SymbolOutput("OperationalEnvelope1", "Operational envelope of the handrail")]

        public AspectDefinition operationalAspect;
        //Centerline Aspect
        [Aspect("Centerline", "Centerline Aspect of Handrail", AspectID.Centerline)]

        public AspectDefinition centerlineAspect;
        #endregion

        #region Construction of outputs of all aspects

        /// <summary>
        /// Creates collection of output for SimplePhysical and Operational Aspect Definitions. 
        /// </summary>
        protected override void ConstructOutputs()
        {
            // Define handrail Property Values
            InitializeHandrailPropertyValues();

            // Validate Handrail Property Values
            ValidateHandrailPropertyValues();

            // Proceed only if to-do messages are not created while validating handrail property values.
            if (!IsErrorToDoMessageCreated)
            {
                // Set handrail property values that are based on other property values.
                SetHandrailPropertyValues();

                // Proceed only if to-do messages are not created while setting handrail property values.
                if (!IsErrorToDoMessageCreated)
                {
                    // Create the handrail outputs of all aspects
                    CreateOutputs();
                }
            }
        }

        #endregion

        #region Private Functions and Methods

        /// <summary>
        /// Initializes the handrail property values.
        /// </summary>
        private void InitializeHandrailPropertyValues()
        {
            // Initizalize the handrail property values
            handrailPropertyValues = new HandRailPropertyValues();

            // If handrail is null, it means it is getting created in S3DHost using sketch path. So set the property valus from Input Definition on the handrail..
            if (handRail == null)
            {
                handrailPropertyValues.sketchPath = (Sketch3D)sketch3d.Value;
                handrailPropertyValues.height = height.Value;
                handrailPropertyValues.toprailSectionAngle = toprailSectionAngle.Value;
                handrailPropertyValues.topOfMidrailDimension = topOfMidrailDimension.Value;
                handrailPropertyValues.midrailSpacing = midrailSpacing.Value;
                handrailPropertyValues.topOfToePlateDimension = topOfToePlateDimension.Value;
                handrailPropertyValues.beginExtensionLength = beginExtensionLength.Value;
                handrailPropertyValues.endExtensionLength = endExtensionLength.Value;
                handrailPropertyValues.segmentMaximumSpacing = segmentMaximumSpacing.Value;
                handrailPropertyValues.slopedSegmentMaximumSpacing = slopedSegmentMaximumSpacing.Value;
                handrailPropertyValues.postSectionAngle = postSectionAngle.Value;
                handrailPropertyValues.midrailSectionAngle = midrailSectionAngle.Value;
                handrailPropertyValues.toePlateSectionAngle = toePlateSectionAngle.Value;

                handrailPropertyValues.horizontalOffsetType = Convert.ToInt32(horizontalOffsetPosition.Value);
                handrailPropertyValues.beginTreatmentType = Convert.ToInt32(beginTreatmentType.Value);
                handrailPropertyValues.endTreatmentType = Convert.ToInt32(endTreatmentType.Value);
                handrailPropertyValues.toprailSectionCP = Convert.ToInt32(topRailSectionCardinalPoint.Value);
                handrailPropertyValues.numberOfMidrails = Convert.ToInt32(numberOfMidrails.Value);
                handrailPropertyValues.orientationValue = (HandrailPostOrientation)handrailOrientation.Value;
                handrailPropertyValues.postSectionCP = Convert.ToInt32(postSectionCardinalPoint.Value);
                handrailPropertyValues.midrailSectionCP = Convert.ToInt32(midrailSectionCardinalPoint.Value);
                handrailPropertyValues.toePlateSectionCP = Convert.ToInt32(toePlateSectionCardinalPoint.Value);
                handrailPropertyValues.postConnectionType = Convert.ToInt32(postConnectionType.Value);
                handrailPropertyValues.horizontalOffsetDimension = Convert.ToInt32(horizontalOffsetDimension.Value);

                handrailPropertyValues.toprailSectionName = Convert.ToString(toprailSectionName.Value);
                handrailPropertyValues.toprailSectionReferenceStandard = Convert.ToString(toprailSectionReferenceStandard.Value);
                handrailPropertyValues.midrailSectionName = Convert.ToString(midrailSectionName.Value);
                handrailPropertyValues.midrailSectionReferenceStandard = Convert.ToString(midrailSectionReferenceStandard.Value);
                handrailPropertyValues.toePlateSectionName = Convert.ToString(toePlateSectionName.Value);
                handrailPropertyValues.toePlateSectionReferenceStandard = Convert.ToString(toePlateSectionReferenceStandard.Value);
                handrailPropertyValues.postSectionName = Convert.ToString(postSectionName.Value);
                handrailPropertyValues.postSectionReferenceStandard = Convert.ToString(postSectionReferenceStandard.Value);

                handrailPropertyValues.isWithToePlate = Convert.ToBoolean(withToePlate.Value);
                handrailPropertyValues.isPostAtEveryTurn = Convert.ToBoolean(withPostAtTurn.Value);
                handrailPropertyValues.connection = OccurrenceConnection;

                handrailPropertyValues.material = primaryMaterial.Value;
                handrailPropertyValues.grade = primaryMaterialGrade.Value;
            }
            else
            // Get the property values from the handrail object.
            {
                handrailPropertyValues.sketchPath = handRail.Path;
                handrailPropertyValues.height = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.Height);
                handrailPropertyValues.toprailSectionAngle = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.TopRailSectionAngle);
                handrailPropertyValues.topOfMidrailDimension = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.TopOfMidrailDimension);
                handrailPropertyValues.midrailSpacing = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.MidRailSpacing);
                handrailPropertyValues.topOfToePlateDimension = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.TopOfToePlateDimension);
                handrailPropertyValues.beginExtensionLength = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.BeginExtensionLength);
                handrailPropertyValues.endExtensionLength = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.EndExtensionLength);
                handrailPropertyValues.segmentMaximumSpacing = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.SegmentMaxSpacing);
                handrailPropertyValues.slopedSegmentMaximumSpacing = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.SlopedSegmentMaxSpacing);
                handrailPropertyValues.postSectionAngle = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.PostSectionAngle);
                handrailPropertyValues.midrailSectionAngle = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.MidRailSectionAngle);
                handrailPropertyValues.toePlateSectionAngle = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToePlateSectionAngle);

                handrailPropertyValues.horizontalOffsetType = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJHandrailProperties, SPSSymbolConstants.HorizontalOffset);
                handrailPropertyValues.beginTreatmentType = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJHandrailProperties, SPSSymbolConstants.BeginTreatmentType);
                handrailPropertyValues.endTreatmentType = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJHandrailProperties, SPSSymbolConstants.EndTreatmentType);
                handrailPropertyValues.toprailSectionCP = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToprailSectionCP);
                handrailPropertyValues.numberOfMidrails = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.NoOfMidRails);
                handrailPropertyValues.orientationValue = (HandrailPostOrientation)StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.HandrailOrientation);
                handrailPropertyValues.postSectionCP = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.PostSectionCP);
                handrailPropertyValues.midrailSectionCP = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.MidRailSectionCP);
                handrailPropertyValues.toePlateSectionCP = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToePlateSectionCP);
                handrailPropertyValues.postConnectionType = StructHelper.GetIntProperty(handRail, SPSSymbolConstants.IJHandrailProperties, SPSSymbolConstants.PostConnectionType);
                handrailPropertyValues.horizontalOffsetDimension = StructHelper.GetDoubleProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.HorizontalOffsetDimension);

                handrailPropertyValues.toprailSectionName = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToprailSectionName);
                handrailPropertyValues.toprailSectionReferenceStandard = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToprailSectionRefStandard);
                handrailPropertyValues.midrailSectionName = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.MidrailSectionName);
                handrailPropertyValues.midrailSectionReferenceStandard = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.MidrailSectionRefStandard);
                handrailPropertyValues.toePlateSectionName = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToePlateSectionName);
                handrailPropertyValues.toePlateSectionReferenceStandard = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToePlateSectionRefStandard);
                handrailPropertyValues.postSectionName = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.PostSectionName);
                handrailPropertyValues.postSectionReferenceStandard = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.PostSectionRefStandard);

                handrailPropertyValues.isWithToePlate = StructHelper.GetBoolProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.WithToePlate);
                handrailPropertyValues.isPostAtEveryTurn = StructHelper.GetBoolProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.WithPostAtTurn);
                handrailPropertyValues.connection = handRail.DBConnection;

                handrailPropertyValues.material = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.PrimaryMaterial);
                handrailPropertyValues.grade = StructHelper.GetStringProperty(handRail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.PrimaryMaterialGrade);
            }
        }

        /// <summary>
        /// Validates the handrail property values.
        /// </summary>
        private void ValidateHandrailPropertyValues()
        {
            // --------------- Validate Orientation ------------------ //
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.HandrailOrientation, SPSSymbolConstants.REFDAT, (int)handrailPropertyValues.orientationValue)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_HANDRAIL_ORIENTATION,
                                                      String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrOrientationCodeListValue,
                                                     "Error while validating handrail orientation in HandrailTypeA symbol as {0} doesnt not exist in {1} table.Check your custom code or contact S3D support."), ((int)handrailPropertyValues.orientationValue).ToString(), SPSSymbolConstants.HandrailOrientation));

                return;
            }

            // -------------------- Valdate Horizontal Offset Type -------------- //
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.HandrailOffset, SPSSymbolConstants.REFDAT, handrailPropertyValues.horizontalOffsetType)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_HORIZONTAL_OFFSET,
                                                      String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrHandrailOffsetCodeListValue,
                                                     "Error while validating handrail offsetType in HandrailTypeA symbol as {0} doesnt not exist in {1} table. Check your custom code or contact S3D support."), (handrailPropertyValues.horizontalOffsetType).ToString(), SPSSymbolConstants.HandrailOffset));

                return;
            }

            // -------------- Validate Treatment Types ---------------- //
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.HandrailEndTreatment, SPSSymbolConstants.REFDAT, handrailPropertyValues.beginTreatmentType)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_BEGIN_TREATMENT_TYPE,
                                                      String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrHandrailTreatmentCodeListValue,
                                                     "Error while validating handrail Treatment Type in HandrailTypeA symbol as {0} doesnt not exist in {1} table. Check your custom code or contact S3D support."), (handrailPropertyValues.beginTreatmentType).ToString(), SPSSymbolConstants.HandrailEndTreatment));
                return;
            }

            // -------------- Validate ConnectionType -----------------//
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.HandrailConnectionType, SPSSymbolConstants.REFDAT, handrailPropertyValues.postConnectionType)))
            {

                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_POST_CONNECTION_TYPE,
                                                      String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrHandrailConnTypeCodeListValue,
                                                     "Error while validating handrail Connection Type in HandrailTypeA symbol as {0} doesnt not exist in {1} table. Check your custom code or contact S3D support."), (handrailPropertyValues.postConnectionType).ToString(), SPSSymbolConstants.HandrailConnectionType));
                return;
            }

            // -------------- Validate Toprail Section Cardinal Point ------------ //
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.CrossSectionCardinalPoints, SPSSymbolConstants.CMNSCH, handrailPropertyValues.toprailSectionCP)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_TOPRAIL_SECTION_CP,
                                                      String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrCrossSectionCPCodeListValue,
                                                     "Error while validating Toprail section cardinal point in HandrailTypeA symbol as {0} doesnt not exist in {1} table.Check your custom code or contact S3D support."), (handrailPropertyValues.toprailSectionCP).ToString(), SPSSymbolConstants.CrossSectionCardinalPoints));
                return;
            }

            // -------------- Validate Midrail Section Cardinal Point ------------ //
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.CrossSectionCardinalPoints, SPSSymbolConstants.CMNSCH, handrailPropertyValues.midrailSectionCP)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_MIDRAIL_SECTION_CP,
                                                     String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrMidrailCrossSectionCPCodeListValue,
                                                    "Error while validating Midrail section cardinal point in HandrailTypeA symbol as {0} doesnt not exist in {1} table. Check your custom code or contact S3D support."), (handrailPropertyValues.midrailSectionCP).ToString(), SPSSymbolConstants.CrossSectionCardinalPoints));
                return;
            }

            // -------------- Validate Toeplate Section Cardinal Point ------------ //
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.CrossSectionCardinalPoints, SPSSymbolConstants.CMNSCH, handrailPropertyValues.toePlateSectionCP)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_TOEPLATE_SECTION_CP,
                                                    String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrToePlateCrossSectionCPCodeListValue,
                                                   "Error while validating Toeplate section cardinal point in HandrailTypeA symbol as {0} doesnt not exist in {1} table. Check your custom code or contact S3D support."), (handrailPropertyValues.toePlateSectionCP).ToString(), SPSSymbolConstants.CrossSectionCardinalPoints));
                return;
            }

            // -------------- Validate Post Section Cardinal Point ------------ //
            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.CrossSectionCardinalPoints, SPSSymbolConstants.CMNSCH, handrailPropertyValues.postSectionCP)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_POST_SECTION_CP,
                                                    String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrToePostSectionCPCodeListValue,
                                                   "Error while validating Post section cardinal point in HandrailTypeA symbol as {0} doesnt not exist in {1} table. Check your custom code or contact S3D support."), (handrailPropertyValues.postSectionCP).ToString(), SPSSymbolConstants.CrossSectionCardinalPoints));
                return;
            }

            // -------------- Validate Toprail Section ReferenceStandard ------------- //
            if (string.IsNullOrEmpty(handrailPropertyValues.toprailSectionReferenceStandard) | string.IsNullOrEmpty(handrailPropertyValues.toprailSectionName))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrMissingToprailSectionProperty,
                                                        "Error while validating Toprail section properties in HandrailTypeA symbol as {0} or {1} or both of them are missing. Please check catalog or contact S3D support."), handrailPropertyValues.toprailSectionReferenceStandard, handrailPropertyValues.toprailSectionName));
                return;
            }

            // ------------- Validate Midrail Section ReferenceStandard ----------------//
            if (string.IsNullOrEmpty(handrailPropertyValues.midrailSectionReferenceStandard) | string.IsNullOrEmpty(handrailPropertyValues.midrailSectionName))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrMissingMidrailSectionProperty,
                                                        "Error while validating Midrail section properties in HandrailTypeA symbol as {0} or {1} or both of them are missing. Please check catalog or contact S3D support."), handrailPropertyValues.midrailSectionReferenceStandard, handrailPropertyValues.midrailSectionName));
                return;
            }

            // ------------- Validate Post Section ReferenceStandard --------------------//
            if (string.IsNullOrEmpty(handrailPropertyValues.postSectionReferenceStandard) | string.IsNullOrEmpty(handrailPropertyValues.postSectionName))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrMissingPostSectionProperty,
                                                        "Error while validating Post section properties in HandrailTypeA symbol as {0} or {1} or both of them are missing. Please check catalog or contact S3D support."), handrailPropertyValues.postSectionReferenceStandard, handrailPropertyValues.postSectionName));
                return;
            }

            // ------------- Validate ToePlate Section ReferenceStandard ----------------//
            if (string.IsNullOrEmpty(handrailPropertyValues.toePlateSectionReferenceStandard) | string.IsNullOrEmpty(handrailPropertyValues.toePlateSectionName))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrMissingToePlateSectionProperty,
                                                         "Error while validating ToePlate section properties in HandrailTypeA symbol as {0} or {1} or both of them are missing. Please check catalog or contact S3D support."), handrailPropertyValues.toePlateSectionReferenceStandard, handrailPropertyValues.toePlateSectionName));
                return;
            }
        }

        /// <summary>
        /// Sets the handrail property values.
        /// </summary>
        private void SetHandrailPropertyValues()
        {
            // Check HandRail Orientation
            if (handrailPropertyValues.orientationValue != HandrailPostOrientation.Vertical
            && handrailPropertyValues.orientationValue != HandrailPostOrientation.Perpendicular)
            {
                //unsupported handrail orientation type
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                       String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrInvalidOrientationType,
                      "Error while checking the handrail orientation in HandrailTypeA symbol as orientationValue: {0} is equal to neither Vertical: {1} nor Perpendicular: {2} values and this orientation is not supported. Check your custom code or contact S3D support."),
                        handrailPropertyValues.orientationValue, HandrailPostOrientation.Vertical, HandrailPostOrientation.Perpendicular));
                return;
                //stop evaluating
            }

            // ------ Set the trace curve/horizontal offset path -------- //
            // Get the handrail complex string from the sketch object
            ComplexString3d handrailSketchPath = (ComplexString3d)handrailPropertyValues.sketchPath.ComplexString;

            if (handrailPropertyValues.horizontalOffsetType == SPSSymbolConstants.LEFT_OFFSET)
            {
                handrailPropertyValues.handrailOffsetPath = (ComplexString3d)handrailSketchPath.GetOffsetCurve(handrailPropertyValues.horizontalOffsetDimension, Common.Middle.OffsetDirection.Coplanar);
            }
            else if (handrailPropertyValues.horizontalOffsetType == SPSSymbolConstants.RIGHT_OFFSET)
            {
                handrailPropertyValues.handrailOffsetPath = (ComplexString3d)handrailSketchPath.GetOffsetCurve(-handrailPropertyValues.horizontalOffsetDimension, Common.Middle.OffsetDirection.Coplanar);
            }
            else
            {
                handrailPropertyValues.horizontalOffsetDimension = 0.0;
                handrailPropertyValues.handrailOffsetPath = handrailSketchPath;
            }

            // Get the handrail sketch path curves
            handrailSketchPath.GetCurves(out handrailPropertyValues.sketchPathCurves);

            // Use actual orientation only if there is no arc in handrail path
            if (handrailPropertyValues.handrailOffsetPath.HasCurveGeometryType(GeometryType.Arc3d))
            {
                handrailPropertyValues.orientationValue = HandrailPostOrientation.Vertical;
            }

            // Get treatment types
            if (handrailPropertyValues.beginTreatmentType == SPSSymbolConstants.NO_END_TREATMENT)
            {
                handrailPropertyValues.beginExtensionLength = 0.0;
            }

            if (!(StructHelper.IsValidCodeListValue(SPSSymbolConstants.HandrailEndTreatment, SPSSymbolConstants.REFDAT, handrailPropertyValues.endTreatmentType)))
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_END_TREATMENT_TYPE, String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrHandrailTreatmentCodeListValue,
                    "Error while validating Handrail endtreatment type in HandrailTypeA symbol as {0} doesnt not exist in {1} table.Check your custom code or contact S3D support."), (handrailPropertyValues.endTreatmentType).ToString(), SPSSymbolConstants.HandrailEndTreatment));

                return;
            }
            if (handrailPropertyValues.endTreatmentType == SPSSymbolConstants.NO_END_TREATMENT)
            {
                handrailPropertyValues.endExtensionLength = 0.0;
            }

            // Get the top rail cross section
            handrailPropertyValues.topRailCrossSection = base.GetCrossSection(handrailPropertyValues.toprailSectionReferenceStandard,
                                                                            handrailPropertyValues.toprailSectionName);

            // If ToDoList message is created, do not proceed
            if(IsErrorToDoMessageCreated)
            {
                return;
            }

            // Get the mid rail cross section
            handrailPropertyValues.midRailCrossSection = base.GetCrossSection(handrailPropertyValues.midrailSectionReferenceStandard,
                                                                            handrailPropertyValues.midrailSectionName);

            // If ToDoList message is created, do not proceed
            if (IsErrorToDoMessageCreated)
            {
                return;
            }

            // Get the toe plate cross section
            handrailPropertyValues.toePlateCrossSection = base.GetCrossSection(handrailPropertyValues.toePlateSectionReferenceStandard,
                                                                            handrailPropertyValues.toePlateSectionName);

            // If ToDoList message is created, do not proceed
            if (IsErrorToDoMessageCreated)
            {
                return;
            }
            // Get the post cross section
            handrailPropertyValues.postCrossSection = base.GetCrossSection(handrailPropertyValues.postSectionReferenceStandard,
                                                                            handrailPropertyValues.postSectionName);

            // If ToDoList message is created, do not proceed
            if (IsErrorToDoMessageCreated)
            {
                return;
            }
            // Set sweep option to create caps
            handrailPropertyValues.sweepOptions = SweepOptions.CreateCaps;


            // Set toprail section depth and width   
            handrailPropertyValues.toprailSectionHeight = handrailPropertyValues.topRailCrossSection.Depth;
            handrailPropertyValues.toprailSectionWidth = handrailPropertyValues.topRailCrossSection.Width;


            //Get the maximum possible distance covered by the cross-section based on the selected
            //cardinal point and the angle rotated for toprail
            double toprailCenterX = 0.0;
            double toprailCenterY = 0.0;
            handrailPropertyValues.toprailSectionMaximumHeight = HandrailServices.GetMaximumProjectedDimensionForSection(handrailPropertyValues.topRailCrossSection, handrailPropertyValues.toprailSectionCP,
                                                                                handrailPropertyValues.toprailSectionAngle, ref toprailCenterX, ref toprailCenterY);

            //Get the circular treatment offset  
            if (handrailPropertyValues.beginTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT |
                handrailPropertyValues.endTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
            {
                //get the height between path and last midrail
                handrailPropertyValues.lastMidrailHeight = handrailPropertyValues.topOfMidrailDimension;
                for (int i = 1; i <= handrailPropertyValues.numberOfMidrails - 1; i++)
                {
                    if (handrailPropertyValues.midrailSpacing > 0.0 & handrailPropertyValues.lastMidrailHeight > handrailPropertyValues.topOfToePlateDimension)
                    {
                        handrailPropertyValues.lastMidrailHeight = handrailPropertyValues.topOfMidrailDimension - handrailPropertyValues.midrailSpacing;
                    }
                }

                // PostSectionDepth & MidRailSectionDepth  considered for consistency with similar calculations in CreateCirEndTreatment()
                handrailPropertyValues.circularTreatmentOffset = (handrailPropertyValues.height - handrailPropertyValues.lastMidrailHeight - handrailPropertyValues.toprailSectionMaximumHeight) / 6;
                //If the above assumption is not a good one, use the following formula to get offset
                //This value should be consistent with the value of ht1 in CreateCirEndTreatment
                if (handrailPropertyValues.toprailSectionHeight > (handrailPropertyValues.height - handrailPropertyValues.toprailSectionMaximumHeight - handrailPropertyValues.lastMidrailHeight))
                {
                    //in this case, circular treatment would be 1/2 of the space between toprail and midrail
                    handrailPropertyValues.circularTreatmentOffset = (handrailPropertyValues.height - handrailPropertyValues.toprailSectionMaximumHeight - handrailPropertyValues.lastMidrailHeight) / 2;
                }
                else if (handrailPropertyValues.circularTreatmentOffset < handrailPropertyValues.toprailSectionHeight / 2)
                {
                    //consider toprail height and 1/4th of the space between toprail and midrail
                    handrailPropertyValues.circularTreatmentOffset = handrailPropertyValues.toprailSectionHeight / 4 + (handrailPropertyValues.height - handrailPropertyValues.toprailSectionMaximumHeight - handrailPropertyValues.lastMidrailHeight) / 4;
        }
                }
            }

        ///// <summary>
        ///// Determines whether an error or to-do message is created.
        ///// </summary>
        ///// <returns>Boolean indicating whether a to-do message is created.</returns>
        //private bool IsErrorToDoMessageCreated()
        //{
        //    // Check if to-do message is created while validating handrail property values.
        //    if (base.ToDoListMessage != null)
        //    {
        //        // Return true if it is an error message.
        //        if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
        //        {
        //            return true;
        //        }
        //    }

        //    // No to-do message created, hence return false;
        //    return false;
        //}

        /// <summary>
        /// Creates the handrail outputs for all aspects
        /// </summary>
        private void CreateOutputs()
        {
            // Create curves
            ComplexString3d toprailPath;
            ComplexString3d topRailCurve;
            Collection<ComplexString3d> midRailCurves;
            ComplexString3d toePlateCurve = null;
            //Collection<ComplexString3d> postCurves;
            Dictionary<string, Curve3d> endTreatmentCurves;

            // -------------- Get Rail Offsets ----------------- //
            // Get begin and end offset values for top and bottom rails.
            double toprailBeginOffset, toprailEndOffset, bottomRailBeginOffset, bottomRailEndOffset;
            GetRailOffsets(out toprailBeginOffset, out toprailEndOffset, out bottomRailBeginOffset, out bottomRailEndOffset);

            // -------------- Get TopRail Curves --------------- //
            try
            {
                // Get the toprail curve
                topRailCurve = HandrailServices.CreateToprailCurve(handrailPropertyValues.handrailOffsetPath, handrailPropertyValues.height - handrailPropertyValues.toprailSectionMaximumHeight,
                                                                handrailPropertyValues.orientationValue, toprailBeginOffset, toprailEndOffset,
                                                                handrailPropertyValues.beginTreatmentType, handrailPropertyValues.endTreatmentType);
            }
            catch (Exception ex)
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, ex.Message);
                return;
                //stop evaluating
            }

            // -------------- Get MidRail Curves --------------- //
            double lowestMidrailHeight = handrailPropertyValues.topOfMidrailDimension;
            try
            {
                // Get midrail curves
                midRailCurves = HandrailServices.CreateMidrailCurves(handrailPropertyValues.handrailOffsetPath, handrailPropertyValues.orientationValue,
                                                                handrailPropertyValues.numberOfMidrails, handrailPropertyValues.midrailSpacing, handrailPropertyValues.topOfMidrailDimension,
                                                                handrailPropertyValues.topOfToePlateDimension, ref lowestMidrailHeight,
                                                                handrailPropertyValues.lastMidrailHeight, handrailPropertyValues.beginTreatmentType,
                                                                handrailPropertyValues.endTreatmentType, bottomRailBeginOffset, bottomRailEndOffset);
            }
            catch (Exception ex)
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, ex.Message);
                return;
                //stop evaluating
            }

            //If orientation is "Perpendicular" then get appropriate path segment so that posts
            //are placed at proper location. Here we are first getting path at TopRail position with perpendicular orientation
            //then project it to original handrail path to get appropriate start-end points of each segment                    
            if (handrailPropertyValues.orientationValue != SPSSymbolConstants.VERTICAL_ORIENTATION)
            {
                //offsets normally the handrail path with height - toprailSectionMaximumHeight (half of the top rail section height).
                toprailPath = (ComplexString3d)handrailPropertyValues.handrailOffsetPath.GetOffsetCurve(handrailPropertyValues.height - handrailPropertyValues.toprailSectionHeight / 2, Common.Middle.OffsetDirection.Normal);
            }

            double postHeight = handrailPropertyValues.height - handrailPropertyValues.toprailSectionHeight / 2;
            bool isMirror = false;
            Collection<ICurve> curves = null;
            handrailPropertyValues.handrailOffsetPath.GetCurves(out curves);

            // Get all posts
            Collection<HandrailPost> posts = CreateHandrailPosts(isMirror);
            // Initialize collection of end treatments
            endTreatmentCurves = new Dictionary<string, Curve3d>();

            // Initialize a dictionary of post with its corresponding IsMirror property
            Dictionary<ComplexString3d, bool> postsWithIsMirror = new Dictionary<ComplexString3d, bool>();

            // currentPostCurves is created temporarily that will be added to dictionary containing post with its corresponding mirror value
            Collection<ICurve> currentPostCurves = new Collection<ICurve>();
            foreach (HandrailPost post in posts)
            {
                bool tempMirror = false;
                if (post.PostIndex == posts.Count - 1)
                {
                    tempMirror = !isMirror;
                }
                else
                {
                    tempMirror = isMirror;
                }
                // Check if this post type is end treatment
                if (post.PostType == SPSSymbolConstants.NO_END_TREATMENT)
                {
                    Line3d line = new Line3d(post.BasePosition, post.PostDirection, postHeight);
                    currentPostCurves.Add(line);
                    // Add the line as complex string with its IsMirror value
                    postsWithIsMirror.Add(new ComplexString3d(currentPostCurves), tempMirror);
                    currentPostCurves.Clear();
                }
            }

            // If first and last post are end treatments, add them to endTreatmentCurves collection
            for (int postIndex = 0; postIndex <= posts.Count - 1; postIndex = postIndex + posts.Count - 1)
            {
                HandrailPost post = posts[postIndex];
                bool tempMirror = false;
                if (post.PostIndex == posts.Count - 1)
                {
                    tempMirror = !isMirror;
                }
                else
                {
                    tempMirror = isMirror;
                }

                try
                {
                    if (post.PostType != SPSSymbolConstants.NO_END_TREATMENT)
                    {
                        if (posts[postIndex].PostIndex > 0)
                        {
                            // End treatment case
                            endTreatmentCurves.Add(HREndTreatment, CreateTreatment(post, curves, postHeight, lowestMidrailHeight, isMirror));
                        }
                        else
                        {
                            // Begin treatment case
                            endTreatmentCurves.Add(HRBeginTreatment, CreateTreatment(post, curves, postHeight, lowestMidrailHeight, isMirror));
                        }
                    }
                }
                catch (Exception ex)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, ex.Message);
                    return;
                    //stop evaluating
                }
            }
            #region Get ToePlate Curves
            if (handrailPropertyValues.isWithToePlate)
            {
                try
                {
                    // Get toeplate curve
                    int toePlateOrientation = (int)handrailPropertyValues.orientationValue;
                    toePlateCurve = HandrailServices.CreateToePlate(handrailPropertyValues.handrailOffsetPath, handrailPropertyValues.topOfToePlateDimension, toePlateOrientation);
                }
                catch (Exception ex)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, ex.Message);
                    return;
                    //stop evaluating
                }
            }
            #endregion

            // Set Material object
            CatalogStructHelper catalogStructHelper = new CatalogStructHelper(handrailPropertyValues.connection);
            Material material = catalogStructHelper.GetMaterial(handrailPropertyValues.material, handrailPropertyValues.grade);

            // If handrail is null, it means it is getting created in S3DHost using sketch path; so create handrail from curves.
            if (handRail == null)
            {
                try
                {
                    #region Construction of SimplePhysical and Centerline Aspects
                    //The code within the loop, executed twice for both SimplePhysical and Centerline Aspect
                    AspectID currentAspectID = AspectID.SimplePhysical;
                    AspectDefinition currentAspect = simplePhysicalAspect;

                    for (int loopCounter = 1; loopCounter <= 2; loopCounter++)
                        {
                        if (loopCounter == 2)
                        {
                            // Set Centerline as aspect
                            currentAspectID = AspectID.Centerline;
                            currentAspect = centerlineAspect;
                        }

                        //add handrail path, physical path (offset curve) and not logical (sketched path) as one of the outputs
                        if (currentAspectID == AspectID.SimplePhysical)
                        {
                            base.AddPathAsOutputOfHandrail(handrailPropertyValues.connection, handrailPropertyValues.sketchPathCurves, ref simplePhysicalAspect);
                        }

                        // Add Toprails output
                        Collection<ISurface> isurfaces = new Collection<ISurface>();
                        int outputCount = 0;
                        AddHandrailOutputForGivenCurve(currentAspect, topRailCurve, handrailPropertyValues.topRailCrossSection, handrailPropertyValues.toprailSectionCP,
                                                       handrailPropertyValues.toprailSectionAngle, handrailPropertyValues.sweepOptions, (int) HandrailHelper.HandrailMemberType.Toprail,
                                                       ref outputCount, material, false, "TopRail");

                        // Add Midrails output
                        outputCount = 0;
                        foreach (ComplexString3d midRailCurve in midRailCurves)
                        {
                            AddHandrailOutputForGivenCurve(currentAspect, midRailCurve, handrailPropertyValues.midRailCrossSection, handrailPropertyValues.midrailSectionCP,
                                                            handrailPropertyValues.midrailSectionAngle, handrailPropertyValues.sweepOptions, (int)HandrailHelper.HandrailMemberType.Midrail,
                                                            ref outputCount, material, false, "MidRail");
                        }
                        
                        // Add ToePlates output.
                        outputCount = 0;
                        if (handrailPropertyValues.isWithToePlate && toePlateCurve != null)
                        {
                            AddHandrailOutputForGivenCurve(currentAspect, toePlateCurve, handrailPropertyValues.toePlateCrossSection, handrailPropertyValues.toePlateSectionCP,
                                                            handrailPropertyValues.toePlateSectionAngle, handrailPropertyValues.sweepOptions, (int)HandrailHelper.HandrailMemberType.ToePlate,
                                                            ref outputCount, material, false, "ToePlate");
                        }
                        
                        // Add Posts output
                        outputCount = 0;
                        Dictionary<ComplexString3d, bool>.KeyCollection postsWithIsMirrorKeys = postsWithIsMirror.Keys;
                        foreach (ComplexString3d postLine in postsWithIsMirrorKeys)
                        {
                            AddHandrailOutputForGivenCurve(currentAspect, postLine, handrailPropertyValues.postCrossSection, handrailPropertyValues.postSectionCP,
                                                            handrailPropertyValues.postSectionAngle, handrailPropertyValues.sweepOptions, (int)HandrailHelper.HandrailMemberType.Post,
                                                            ref outputCount, material, false, "HRPost");
                        }
                        
                        // Add End Treatments output
                        outputCount = 0;
                        Dictionary<string, Curve3d>.KeyCollection endTreatmentCurvesKeys = endTreatmentCurves.Keys;
                        foreach (string treatmentName in endTreatmentCurvesKeys)
                        {

                            AddHandrailOutputForGivenCurve(currentAspect, endTreatmentCurves[treatmentName], handrailPropertyValues.topRailCrossSection, handrailPropertyValues.toprailSectionCP,
                                                            handrailPropertyValues.toprailSectionAngle, handrailPropertyValues.sweepOptions, (int)HandrailHelper.HandrailMemberType.Post,
                                                            ref outputCount, material, false, treatmentName);
                        }
                    }
                    #endregion

                    #region Construction of Operational Aspect
                    // Getting complex string curve at toprail location. This can be used to create ruled surface between this curve and handrail path. 
                    // Also can be used for creating cylinder along toprail
                    OffsetDirection offsetDirection = default(OffsetDirection);
                    if (handrailPropertyValues.orientationValue == HandrailPostOrientation.Vertical)
                    {
                        offsetDirection = OffsetDirection.Vertical;
                        //offsets vertically the handrail path with an offset equal to height - deltaHeight (distance between top edge to the current CP)
                        toprailPath = (ComplexString3d)handrailPropertyValues.handrailOffsetPath.GetOffsetCurve(new Vector(0.0, 0.0, handrailPropertyValues.height - handrailPropertyValues.toprailSectionMaximumHeight));
                    }
                    else
                    {
                        //offsets normally the handrail path with an offset equal to height - deltaHeight (distance between top edge to the current CP)
                        offsetDirection = OffsetDirection.Normal;
                        //offsets on the curve plane
                        toprailPath = (ComplexString3d)handrailPropertyValues.handrailOffsetPath.GetOffsetCurve(handrailPropertyValues.height - handrailPropertyValues.toprailSectionMaximumHeight, offsetDirection);
                    }

                    //Creating ruled surface between handrail path and curve along toprail
                    Ruled3d surface = new Ruled3d(toprailPath, handrailPropertyValues.handrailOffsetPath, true);
                    operationalAspect.Outputs.Add("Handrail surface", surface);

                    CrossSectionServices crosssectionServices = new CrossSectionServices();
                    Position pathStartPosition = toprailPath.PointAtDistanceAlong(0.0);
                    Vector tangent = toprailPath.TangentAtPoint(pathStartPosition);
                    //Create a circle and sweep it through the whole sketch path.
                    double clearance = 3 * 0.0254; // 3 Inch clearance
                    Circle3d circle = crosssectionServices.GetBoundedCircleFromCrossSection(handrailPropertyValues.topRailCrossSection, handrailPropertyValues.toprailSectionCP, pathStartPosition, tangent, clearance);

                    //Cover the toprail by a cylinder with sweeping circle along curve at toprail location. 
                    Collection<Surface3d> surfaces = Surface3d.GetSweepSurfacesFromCurve(toprailPath, circle, SurfaceSweepOptions.CreateCaps);
                    if ((surfaces != null))
                    {
                        for (int index = 0; index <= surfaces.Count - 1; index++)
                        {
                            operationalAspect.Outputs.Add("Toprail cylinder" + index.ToString(), surfaces[index]);
                        }
                    }

                    #endregion

                    #region Set weight COG
                    //Now set the weight COG to the aspect.
                    // This call should also set it on the BO in case EvaluateCOG is not called during placement
                    base.SetWeightCOG(simplePhysicalAspect);

                    // For Handrail on members, EvaluateWeightCG is not called at placement. Hence need to set weightCG here.
                    SetWeightCOG();

                    #endregion
                }
                catch (Exception)
                {
                    //Check if ToDoListMessgae already created, if not, create a todo record with generic failure message.
                    if (base.ToDoListMessage == null)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrConstructOutputs, "Error in constructing outputs for handrail in handrailTypeA symbol.Check your custom code or contact S3D support."));
                    }
                }
            }
            // Drop to members
            else
            {
                int memberTypeIndex;

                // Create Toprails
                CreateMembersForGivenCurve(handRail, topRailCurve, 1, handrailPropertyValues.topRailCrossSection, handrailPropertyValues.toprailSectionCP, 
                                    handrailPropertyValues.toprailSectionAngle, material, MemberType.TopRail, false);
                
                // Create Midrails
                memberTypeIndex = 1;
                foreach (ComplexString3d midRailCurve in midRailCurves)
                {
                    // Create members for current midrail curve
                    CreateMembersForGivenCurve(handRail, midRailCurve, memberTypeIndex, handrailPropertyValues.midRailCrossSection, handrailPropertyValues.midrailSectionCP,
                                        handrailPropertyValues.midrailSectionAngle, material, MemberType.MidRail, false);
                    // Get the next member type index
                    memberTypeIndex++;
                }

                // Create ToePlates 
                if (handrailPropertyValues.isWithToePlate && toePlateCurve != null)
                {
                    CreateMembersForGivenCurve(handRail, toePlateCurve, 1, handrailPropertyValues.toePlateCrossSection, handrailPropertyValues.toePlateSectionCP,
                                        handrailPropertyValues.toePlateSectionAngle, material, MemberType.ToePlate, false);
                }

                // Create Posts 
                foreach (ComplexString3d postCurve in postsWithIsMirror.Keys)
                {
                    CreateMembersForGivenCurve(handRail, postCurve, 1, handrailPropertyValues.postCrossSection, handrailPropertyValues.postSectionCP,
                                        handrailPropertyValues.postSectionAngle, material, MemberType.Post, false);
                }

                // Create End Treatments 
                memberTypeIndex = 1;
                foreach (string treatmentName in endTreatmentCurves.Keys)
                {
                    if (string.Compare(treatmentName, HRBeginTreatment, false) == 0)
                    {
                        memberTypeIndex = 1;
                    }
                    else if (string.Compare(treatmentName, HREndTreatment, false) == 0)
                    {
                        memberTypeIndex = 2;
                    }
                    else
                    {
                        // This is unexpected end treatment name. Throw exception.
                        throw new CmnException(String.Format(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrInvalidEndTreatmentType,
                            "End TreatmentType: {0} is other than {1} and {2} while creating End Treatments in HandrailTypeA. Please check the treatmentName"), treatmentName, HRBeginTreatment, HREndTreatment));
                    }
                    // Create members for curve end treatment curve.
                    base.CreateMembers(handRail, memberTypeIndex, (ICurve)endTreatmentCurves[treatmentName], handrailPropertyValues.topRailCrossSection, material,
                                            (int)MemberType.EndTreatment, handrailPropertyValues.toprailSectionCP,
                                                handrailPropertyValues.toprailSectionAngle, false);
                }
            }
        }

        /// <summary>
        /// Creates the members for all the curves in given complex string. 
        /// </summary>
        /// <param name="handRail">The hand rail.</param>
        /// <param name="curve">The curve as complex string.</param>
        /// <param name="memberTypeIndex">Index of the member type.</param>
        /// <param name="crossSection">The cross section.</param>
        /// <param name="sectionCP">The rail section cp.</param>
        /// <param name="sectionAngle">The section angle.</param>
        /// <param name="material">The material.</param>
        /// <param name="memberType">Type of the member.</param>
        /// <param name="isMirror">if set to <c>true</c> [is mirror].</param>
        private void CreateMembersForGivenCurve(HandRail handRail, ComplexString3d curve, int memberTypeIndex, CrossSection crossSection, int sectionCP, double sectionAngle, 
                                        Material material, MemberType memberType, bool isMirror)
        {
            Collection<ICurve> componentCurves = null;
            try
            {
                // Get all curves.
                curve.GetCurves(out componentCurves);

                // Create member system for each curve.
                foreach (ICurve currentComponentCurve in componentCurves)
                {
                    base.CreateMembers(handRail, memberTypeIndex, currentComponentCurve, crossSection, material, (int)memberType, sectionCP,
                                            sectionAngle, isMirror);
                }
            }
            catch(Exception ex)
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrCreateMembersForGivenCurve, "Failed to create members for given curve in HadraiTypeA Symbol:" + ex.ToString()));
            }
        }

        /// <summary>
        /// Adds the handrail output for given curve.
        /// </summary>
        /// <param name="currentAspect">The current aspect.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="crossSection">The cross section.</param>
        /// <param name="sectionCP">The section cp.</param>
        /// <param name="sectionAngle">The section angle.</param>
        /// <param name="sweepOptions">The sweep options.</param>
        /// <param name="memberType">Type of the member.</param>
        /// <param name="material">The material.</param>
        /// <param name="isMirror">if set to <c>true</c> [is mirror].</param>
        /// <param name="outputPrefix">The output prefix.</param>
        /// <exception cref="Ingr.SP3D.Common.Middle.CmnException">Failed to add output for given curve. + ex.ToString()</exception>
        private void AddHandrailOutputForGivenCurve(AspectDefinition currentAspect, Curve3d curve, CrossSection crossSection, int sectionCP, double sectionAngle,
                                                     SweepOptions sweepOptions, int memberType, ref int outputCount, Material material, bool isMirror, string outputPrefix)
        {
            Collection<ISurface> surfaces = new Collection<ISurface>();
            try
            {
                // Add handrail output
                surfaces = HandrailServices.AddHandrailOutput(currentAspect, curve, crossSection, sectionCP, sectionAngle, isMirror,
                                                                outputPrefix, memberType, ref outputCount, sweepOptions);
                
                // Add weight COG if the surface collection in not empty.
                if (surfaces != null)
                {
                    base.AddWeightCOG(curve, crossSection);
                }
            }
            catch (Exception)
            {
                //Check if ToDoListMessgae already created, if not, create a todo record with generic failure message.
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrAddHandrailOutputForGivenCurve,
                        string.Format("Error in creating surfaces for {1} curve while adding handrail outputs in AddHandrailOutputsForGivenCurve in HandrailTypeA Symbol. Please check your custom code or contact S3D support.", outputPrefix)));
                }
            }
        }

        /// <summary>
        /// Gets the top and bottom rail offset values at begin and end.
        /// </summary>
        /// <param name="toprailBeginOffset">The toprail begin offset.</param>
        /// <param name="toprailEndOffset">The toprail end offset.</param>
        /// <param name="bottomRailBeginOffset">The bottom rail begin offset.</param>
        /// <param name="bottomRailEndOffset">The bottom rail end offset.</param>
        private void GetRailOffsets(out double toprailBeginOffset, out double toprailEndOffset, out double bottomRailBeginOffset, out double bottomRailEndOffset)
        {
            toprailBeginOffset = 0.0;
            toprailEndOffset = 0.0;
            bottomRailBeginOffset = 0.0;
            bottomRailEndOffset = 0.0;

            // Get the begin circularTratment offset for toprail.
            toprailBeginOffset = GetCircularTreatmentRadius(handrailPropertyValues.sketchPathCurves[0], true, true,
                                                            handrailPropertyValues.circularTreatmentOffset, 0.0, handrailPropertyValues.orientationValue);

            // Get the end circularTratment offset for toprail.
            toprailEndOffset = GetCircularTreatmentRadius(handrailPropertyValues.sketchPathCurves[handrailPropertyValues.sketchPathCurves.Count - 1],
                                        false, true, handrailPropertyValues.circularTreatmentOffset, 0.0, handrailPropertyValues.orientationValue);

            // Check if there are circular begin treatment and end treatment
            if (handrailPropertyValues.beginTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT ||handrailPropertyValues.endTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
            {
                // Get the begin circularTratment offset for midrail.
                bottomRailBeginOffset = GetCircularTreatmentRadius(handrailPropertyValues.sketchPathCurves[0], true, false,
                                                                    handrailPropertyValues.circularTreatmentOffset, 0.0, handrailPropertyValues.orientationValue);

                // Get the end circularTratment offset for the midrail.
                bottomRailEndOffset = GetCircularTreatmentRadius(handrailPropertyValues.sketchPathCurves[handrailPropertyValues.sketchPathCurves.Count - 1], false, false,
                                                                    handrailPropertyValues.circularTreatmentOffset, 0.0, handrailPropertyValues.orientationValue);
            }
        }

        /// <summary>
        /// Creates the end treatment at given location and returns its curve. It returns complexstring in case of circular end treatment and Line3d in case of 
        /// rectangular end treatment.
        /// </summary>
        /// <param name="post">The post.</param>
        /// <param name="curves">The collection of handrail object's trace curves.</param>
        /// <param name="postHeight">Height of the post.</param>
        /// <param name="lowestMidrailHeight">Height of the lowest midrail.</param>
        /// <param name="isMirror">if set to <c>true</c> [is mirror].</param>
        /// <returns>
        /// Curve pertaining to the treatment.
        /// </returns>
        private Curve3d CreateTreatment(HandrailPost post, Collection<ICurve> curves, double postHeight, double lowestMidrailHeight, bool isMirror)
        {
            Curve3d currentEndTreatment = null;
            Line3d line = default(Line3d);
            ICurve curve = default(ICurve);
            Position postBase = post.BasePosition;
            Vector postDirection = post.PostDirection;

            // Indicates End Treatment
            if (post.PostIndex > 0)
            {
                curve = curves[curves.Count - 1];
            }
            // Indicates Begin Treatment
            else
            {
                curve = curves[0];
            }

            // Create rectangular end treatment
            if (post.PostType == SPSSymbolConstants.RECTANGULAR_END_TREATMENT)
            {
                Position start = new Position();
                Vector direction = new Vector();
                if (post.PostIndex > 0)
                {
                    start.Set(postBase.X + postHeight * postDirection.X, postBase.Y + postHeight * postDirection.Y, postBase.Z + postHeight * postDirection.Z);
                    direction.Set(-postDirection.X, -postDirection.Y, -postDirection.Z);
                    line = new Line3d(start, direction, postHeight - lowestMidrailHeight);
                }
                else
                {
                    start.Set(postBase.X + lowestMidrailHeight * postDirection.X, postBase.Y + lowestMidrailHeight * postDirection.Y, postBase.Z + lowestMidrailHeight * postDirection.Z);
                    direction.Set(postDirection.X, postDirection.Y, postDirection.Z);
                    line = new Line3d(start, direction, postHeight - lowestMidrailHeight);
                }
                //treatmentCurves.Add(line);
                currentEndTreatment = line;
            }
            // Create Circular End Treatment
            else if (post.PostType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
            {
                if (post.PostIndex > 0)
                    isMirror = !isMirror;

                // Get the complex3d curve representing the circular end treatment.
                currentEndTreatment =  CreateCirularEndTreatmentCurve(curve, post.PostIndex == 0,
                                            postBase, lowestMidrailHeight, postHeight, postDirection,
                                            handrailPropertyValues.toprailSectionHeight, handrailPropertyValues.circularTreatmentOffset);
            }
            return currentEndTreatment;
        }

        /// <summary>
        /// Creates all handrail posts.
        /// </summary>
        /// <param name="isMirror">Defines whether it is a post or the treatment needs to be mirrored.</param>
        /// <returns>
        /// Collection of post object which will be used to create post and end treatments.
        /// </returns>
        /// <exception cref="System.ArgumentNullException">curves</exception>
        /// <exception cref="Ingr.SP3D.Common.Middle.CmnException">Handrail not cosntain any input curves, change the handrail path.</exception>
        private Collection<HandrailPost> CreateHandrailPosts(bool isMirror)
        {
            Collection<ICurve> curves = null;
            handrailPropertyValues.handrailOffsetPath.GetCurves(out curves);
            if (curves == null)
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves,
                    "Handrail does not contain any input curves while creating handrail posts in CreateHandrailPosts in HandrailTypeA, change the handrail path.")); ;
            }
            if (curves.Count == 0)
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves,
                    "Handrail does not contain any input curves while creating handrail posts in CreateHandrailPosts in HandrailTypeA, change the handrail path."));
            }

            bool isPostAtEveryTurn = handrailPropertyValues.isPostAtEveryTurn;
            double sectionAngle = 0;
            double startPostSectionAngle = 0.0;
            double endPostSectionAngle = 0.0;
            Collection<bool> isSlopedCurve = new Collection<bool>();
            //Dim geometryServices As New GeometryServices()
            Collection<HandrailPost> posts = new Collection<HandrailPost>();

            //Get the start and end post section angle adjustments for treatments
            if ((handrailPropertyValues.beginExtensionLength > 0 && handrailPropertyValues.beginTreatmentType != SPSSymbolConstants.NO_END_TREATMENT) ||
                (handrailPropertyValues.endExtensionLength > 0 && handrailPropertyValues.endTreatmentType != SPSSymbolConstants.NO_END_TREATMENT))
            {
                startPostSectionAngle = GetPostSectionAngleForEndTreatment(curves[0], true, handrailPropertyValues.toprailSectionAngle);
                endPostSectionAngle = GetPostSectionAngleForEndTreatment(curves[curves.Count - 1], false, handrailPropertyValues.toprailSectionAngle);
            }

            Position startPosition = new Position();
            Position endPosition = new Position();
            Vector postDirection = new Vector();
            Vector tangent = null;
            //, deltaTangent As Vector = Nothing
            Vector segmentVector = new Vector();
            //Dim curveLength As Double, curveDist As Double, curveParameter As Double ' startParameter As Double, endParameter As Double, 
            //Dim startPos As Position = Nothing, endPos As Position = Nothing, 
            Position parameterPos = null;

            //hold first and last point of the complex curve.
            //requires this only for calculating for start and end post positions.
            handrailPropertyValues.handrailOffsetPath.EndPoints(out startPosition, out endPosition);
            for (int i = 0; i <= curves.Count - 1; i++)
            {
                ICurve curve = default(ICurve);
                //get the clearance flags at both ends
                bool clearanceNeededAtStart = false;
                bool clearanceNeededAtEnd = false;
                IsPostClearanceNeededForPathSegment(handrailPropertyValues.handrailOffsetPath, i, out clearanceNeededAtStart, out clearanceNeededAtEnd);
                handrailPropertyValues.isPostAtEveryTurn = isPostAtEveryTurn;

                // get the current curve
                curve = curves[i];
                //get the tangent of at the start point of the curve
                parameterPos = curve.PointAtDistanceAlong(0.0);
                tangent = curve.TangentAtPoint(parameterPos);
                segmentVector = new Vector(tangent);

                if (handrailPropertyValues.orientationValue == HandrailPostOrientation.Perpendicular)
                {
                    //determine post direction _|_er to the curve
                    Vector axis = new Vector(0.0, 0.0, 1.0);
                    Vector vecPDir = axis.Cross(segmentVector);
                    vecPDir.Length = 1;
                    vecPDir = segmentVector.Cross(vecPDir);
                    vecPDir.Length = 1;
                    postDirection = new Vector(vecPDir.X, vecPDir.Y, vecPDir.Z);
                }
                else
                {
                    //use vertical direction
                    postDirection.Set(0.0, 0.0, 1.0);
                }
                double deltaAngle = 0;
                //int postCount = posts.Count;
                //add start treatment position if applicable
                //for first post on a first curve, check if treatment is needed
                if (i == 0 && handrailPropertyValues.beginExtensionLength > 0 && handrailPropertyValues.beginTreatmentType != SPSSymbolConstants.NO_END_TREATMENT)
                {
                    if (handrailPropertyValues.beginTreatmentType == SPSSymbolConstants.RECTANGULAR_END_TREATMENT)
                    {
                        posts.Add(new HandrailPost(startPosition, postDirection, handrailPropertyValues.beginTreatmentType, startPostSectionAngle, posts.Count));
                        //for circluar treatment use toprail section angle
                    }
                    else
                    {
                        deltaAngle = HandrailServices.GetOrientationAngle(handrailPropertyValues.topRailCrossSection, true);
                        posts.Add(new HandrailPost(startPosition, postDirection, handrailPropertyValues.beginTreatmentType, handrailPropertyValues.toprailSectionAngle + deltaAngle, posts.Count));
                        //3 * Math.PI / 2
                    }
                }

                //Add intermediate posts now
                AddIntermediatePosts(curve, i, curves.Count, sectionAngle, clearanceNeededAtStart, clearanceNeededAtEnd, isMirror, postDirection, ref posts);


                //add last post with end treatment position if applicable
                if (i == curves.Count - 1 && handrailPropertyValues.endExtensionLength > 0 && handrailPropertyValues.endTreatmentType != SPSSymbolConstants.NO_END_TREATMENT)
                {
                    if (handrailPropertyValues.endTreatmentType == SPSSymbolConstants.RECTANGULAR_END_TREATMENT)
                    {
                        posts.Add(new HandrailPost(endPosition, postDirection, handrailPropertyValues.endTreatmentType, endPostSectionAngle, posts.Count));
                        //for circluar treatment use toprail section angle 
                    }
                    else
                    {
                        deltaAngle = HandrailServices.GetOrientationAngle(handrailPropertyValues.topRailCrossSection, false);
                        posts.Add(new HandrailPost(endPosition, postDirection, handrailPropertyValues.endTreatmentType, handrailPropertyValues.toprailSectionAngle + deltaAngle, posts.Count));
                    }
                }
            }

            return posts;
        }

        /// <summary>
        /// This method validate handrail for the user keyed in inputs of Height, ToTopOfMidRailDistance, NoOfMidRails and MidRailSpacing.
        /// Need to check these properties in conjunction as they are related to each other.
        /// </summary>
        /// <param name="handrail">Handrail business object which aggregates symbol.</param>
        /// <param name="errorMessage">The error message if validation fails.</param>
        /// <returns>
        /// True if property value validation succeeds.
        /// </returns>
        private bool AreHandrailPropertiesValid(HandRail handrail, ref string errorMessage)
        {
            errorMessage = string.Empty;
            //Get the handrail user attributes to validate
            double handrailHeight = 0.0;
            double midrailDistance = 0.0;
            double midrailSpacing = 0.0;
            long numberOfMidrails = 0;
            try
            {
                handrailHeight = StructHelper.GetDoubleProperty(handrail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.Height);
                midrailDistance = StructHelper.GetDoubleProperty(handrail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.TopOfMidrailDimension);
                midrailSpacing = StructHelper.GetDoubleProperty(handrail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.MidRailSpacing);
                numberOfMidrails = StructHelper.GetLongProperty(handrail, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.NoOfMidRails);
            }
            catch (Exception)
            {
                //attributes might be missing, create todo record
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, handrail.ToString() + HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrUserAttributesMissing,
                    "some of the required user attribute values cannot be obtained from the catalog part while validating handrail in HandrailTypeA Symbol. Check the error log and catalog data."));
                return false;
            }

            //Check whether Top of Midrail distance is more than Handrail Height 
            if (handrailHeight <= midrailDistance)
            {
                errorMessage = "Invalid property value - check user inputs";
                return false;
            }

            //Check whether Number of MidRails are keyed in with a valid MidRail spacing or vice versa 
            if (((numberOfMidrails - 1) * midrailSpacing) >= midrailDistance)
            {
                errorMessage = "Invalid property value - check user inputs";
                return false;
            }

            return true;

        }

        /// <summary>
        /// Adds the intermediate posts to the collection
        /// </summary>
        /// <param name="curve">curve segment</param>
        /// <param name="index">curve index in the handrail path</param>
        /// <param name="curveCount">number of curves in the path</param>
        /// <param name="sectionAngle">section angle</param>
        /// <param name="clearanceNeededAtStart">flag to specify whether clearance needed at start</param>
        /// <param name="clearanceNeededAtEnd">flag to specify whether clearance needed at end</param>
        /// <param name="isMirror">is mirror</param>
        /// <param name="postDirection">post direction vector</param>
        /// <param name="handrailPosts">post collection</param>
        private void AddIntermediatePosts(ICurve curve, int index, int curveCount, double sectionAngle, bool clearanceNeededAtStart,
                                            bool clearanceNeededAtEnd, bool isMirror,
                                            Vector postDirection, ref Collection<HandrailPost> handrailPosts)
        {
            double deltaAngle = Math.PI;
            double postSpacing = 0;

            //get the curve length
            double curveLength = curve.Length;
            //check if the given segment is horizontal or not
            bool isSegmentSloped = !IsSegmentHorizontal(curve);

            // override user input for postAtTurn flag if the given conditions are satisfied.
            if (IsPostAtTurnNeeded(isSegmentSloped, clearanceNeededAtStart, clearanceNeededAtEnd) == true)
            {
                handrailPropertyValues.isPostAtEveryTurn = true;
            }

            //Dim sP As Double = startParameter, eP As Double = endParameter
            double startExtention = 0.0;
            //leave space for begin treatment extension
            if (index == 0 && handrailPropertyValues.beginExtensionLength > 0 && handrailPropertyValues.beginTreatmentType != SPSSymbolConstants.NO_END_TREATMENT)
            {
                curveLength = curveLength - handrailPropertyValues.beginExtensionLength;
                startExtention = handrailPropertyValues.beginExtensionLength;
            }

            //leave space for end treatment extension
            if (index == curveCount - 1 && handrailPropertyValues.endExtensionLength > 0 && handrailPropertyValues.endTreatmentType != SPSSymbolConstants.NO_END_TREATMENT)
            {
                curveLength = curveLength - handrailPropertyValues.endExtensionLength;
            }

            //initialize to middle segment
            SegmentType segmentType = SegmentType.Middle;

            if (index == 0)
            {
                //start curve
                segmentType = SegmentType.Begin;
            }
            else if (index == curveCount - 1)
            {
                //end curve
                segmentType = SegmentType.End;
            }

            //now calculate intermediate posts and post spacing  
            PostsAtTurnInformation postsAtTurnInfo = default(PostsAtTurnInformation);
            Position positionAtDistance = default(Position);
            //TR-236718
            if (curveLength <= Math3d.DistanceTolerance)
            {
                //don’t process if the curve length is less than default extension length as there is no room for placing the intermediate posts
                //this happens only for endcurves of the path whose length are less than the extension lengths
                if (index == curveCount - 1 && handrailPropertyValues.isPostAtEveryTurn)
                {
                    //if it is  last curve,create a post at start if PostAtEveryTurn flag is on and return since intermediate post includes first
                    positionAtDistance = curve.PointAtDistanceAlong(0.0);
                    handrailPosts.Add(new HandrailPost(positionAtDistance, postDirection, SPSSymbolConstants.NO_END_TREATMENT, sectionAngle, handrailPosts.Count));
                }
                return;
            }
            else
            {
                postsAtTurnInfo = GetPostsAtTurnInformation(curveLength, segmentType, isSegmentSloped, clearanceNeededAtStart, clearanceNeededAtEnd, handrailPropertyValues.isPostAtEveryTurn);
            }
            if (postsAtTurnInfo.IntermediatePostsCount > 0)
            {
                postSpacing = curveLength * (1 - postsAtTurnInfo.ClearancePercentBeforeTurn - postsAtTurnInfo.ClearancePercentAfterTurn) / postsAtTurnInfo.IntermediatePostsCount;
            }

            //evaluate post postions for all intermediate posts
            for (int j = 0; j <= postsAtTurnInfo.IntermediatePostsCount; j++)
            {
                //dont process for last post of the each curve, if post at turn is on
                if (j == postsAtTurnInfo.IntermediatePostsCount && index < curveCount - 1 && handrailPropertyValues.isPostAtEveryTurn)
                {
                    break; // TODO: might not be correct. Was : Exit For
                }
                double distanceAlong = 0;
                distanceAlong = j * postSpacing + startExtention + curveLength * postsAtTurnInfo.ClearancePercentAfterTurn;
                //get the section angle after aligning with path
                sectionAngle = GetSectionAngleWithPath(curve, distanceAlong, handrailPropertyValues.postSectionAngle);

                //get the position on the path at given curve paramenter.
                positionAtDistance = curve.PointAtDistanceAlong(distanceAlong);
                //reverse the angle for the last post
                if (index == curveCount - 1 && j == postsAtTurnInfo.IntermediatePostsCount && handrailPropertyValues.endTreatmentType == SPSSymbolConstants.NO_END_TREATMENT)
                {
                    if (isMirror)
                    {
                        sectionAngle = sectionAngle - deltaAngle;
                    }
                    else
                    {
                        sectionAngle = sectionAngle + deltaAngle;
                    }
                }
                handrailPosts.Add(new HandrailPost(positionAtDistance, postDirection, SPSSymbolConstants.NO_END_TREATMENT, sectionAngle, handrailPosts.Count));
            }
        }

        /// <summary>
        /// Returns true if a curved member exists in given collection of components, else returns false.
        /// </summary>
        /// <param name="components">The components.</param>
        /// <returns></returns>
        private bool CurvedRailExists(ReadOnlyDictionary<int, ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>>> components)
        {
            // Process all individual members in all components
            foreach (int key in components.Keys)
            {
                ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>> component = components[key];
                foreach (int componentKey in component.Keys)
                {
                    ReadOnlyCollection<BusinessObject> businessObjects = component[componentKey];
                    foreach (BusinessObject componentBO in businessObjects)
                    {
                        if (componentBO is MemberSystem)
                        {
                            // Return true if member system is curved
                            if (((MemberSystem)componentBO).Curved)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }
        #endregion

        #region ICustomWeightCG Members

        /// <summary>
        /// Evaluates the weight and center of gravity of the handrail part and sets it on the handrail business object.
        /// </summary>
        /// <param name="businessObject">Handrail business object which aggregates symbol.</param>
        public void EvaluateWeightCG(BusinessObject businessObject)
        {
            //Set weight and COG as set during ConstructOutputs
            base.SetWeightCOG();            
        }

        #endregion

        #region Overrides Functions And Methods

        /// <summary>
        /// Makes the property read-only.
        /// </summary>
        /// <param name="interfaceName">Interface name of the property.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <returns></returns>
        public override bool IsPropertyReadOnly(string interfaceName, string propertyName)
        {
            bool result = false;
            switch (propertyName)
            {
                //making following property values as read-only. 
                case SPSSymbolConstants.ToprailSectionRefStandard:
                case SPSSymbolConstants.MidrailSectionRefStandard:
                case SPSSymbolConstants.ToePlateSectionRefStandard:
                case SPSSymbolConstants.PostSectionRefStandard:
                case SPSSymbolConstants.TotalLength:
                    result = true;
                    break;
            }
            return result;
        }

        /// <summary>
        /// Validates the given property.
        /// </summary>
        /// <param name="handrail">Handrail business object which aggregates symbol.</param>
        /// <param name="interfaceName">Interface name of the property.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <param name="propertyValue">The value of the property.</param>
        /// <param name="errorMessage">The error message if validation fails.</param>
        /// <returns>True if property value validation succeeds.</returns>
        public override bool IsPropertyValid(HandRail handRail, string interfaceName, string propertyName, object propertyValue, out string errorMessage)
        {

            //by default set the property value as valid. Override the value later for known checks
            bool isValidPropertyValue = true;
            string errorMessage2 = String.Empty;
            if (propertyValue != null)
            {
                switch (propertyName)
                {
                    //following property value has combo-box to select the proper option, so set the property value as valid. 
                    case SPSSymbolConstants.ToprailSectionName:
                    case SPSSymbolConstants.MidrailSectionName:
                    case SPSSymbolConstants.ToePlateSectionName:
                    case SPSSymbolConstants.PostSectionName:
                    case SPSSymbolConstants.PrimaryMaterial:
                        isValidPropertyValue = true;
                        break;

                    //following property values need to be in between 0-360 degrees 
                    case SPSSymbolConstants.TopRailSectionAngle:
                    case SPSSymbolConstants.MidRailSectionAngle:
                    case SPSSymbolConstants.ToePlateSectionAngle:
                    case SPSSymbolConstants.PostSectionAngle:
                        isValidPropertyValue = ValidationHelper.IsBetween0And360(Convert.ToDouble(propertyValue), ref errorMessage2);
                        break;

                    //following property values must be greater than 0
                    case SPSSymbolConstants.Height:
                    case SPSSymbolConstants.TopOfMidrailDimension:
                    case SPSSymbolConstants.MidRailSpacing:
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(Convert.ToDouble(propertyValue), ref errorMessage2);
                        if (isValidPropertyValue)
                        {
                            isValidPropertyValue = AreHandrailPropertiesValid(handRail, ref errorMessage2);
                        }
                        break;

                    //following property values must be greater than 0
                    case SPSSymbolConstants.NoOfMidRails:
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(Convert.ToInt32(propertyValue), ref errorMessage2);
                        if (isValidPropertyValue)
                        {
                            isValidPropertyValue = AreHandrailPropertiesValid(handRail, ref errorMessage2);
                        }
                        break;

                    //following property values must be greater than 0
                    case SPSSymbolConstants.HorizontalOffsetDimension:
                    case SPSSymbolConstants.SegmentMaxSpacing:
                    case SPSSymbolConstants.SlopedSegmentMaxSpacing:
                    case SPSSymbolConstants.TopOfToePlateDimension:
                    case SPSSymbolConstants.BeginExtensionLength:
                    case SPSSymbolConstants.EndExtensionLength:
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(Convert.ToDouble(propertyValue), ref errorMessage2);
                        break;
                }
            }
            errorMessage = errorMessage2;
            return isValidPropertyValue;
        }

        /// <summary>
        /// Gets the allowable post clearance values at post turn
        /// </summary>
        /// <param name="minimumClearanceAtPostTurn"></param>
        /// <param name="maximumClearanceAtPostTurn"></param>
        /// <remarks></remarks>
        protected override void GetAllowablePostClearanceValues(out double minimumClearanceAtPostTurn, out double maximumClearanceAtPostTurn)
        {
            //currently set these values to be 2Feet for maximum and 9inches for minimum clearance at post turn, 
            maximumClearanceAtPostTurn = 0.6096;
            minimumClearanceAtPostTurn = 0.2286;
        }

        /// <summary>
        /// Gets the allowable spacing between posts
        /// </summary>
        /// <param name="isSegmentSloped">flag to specify whether the segment is sloped.</param>
        /// <remarks></remarks>
        protected override double GetMaximumAllowableSpacingBetweenPosts(bool isSegmentSloped)
        {
            //set the spacing based on the segment orinetation.
            if (isSegmentSloped)
            {
                if (slopedSegmentMaximumSpacing == null)
                {
                    return 1.829;
                    //default value 6 feet
                }
                else
                {
                    return slopedSegmentMaximumSpacing.Value;
                }
            }
            else
            {
                if (segmentMaximumSpacing == null)
                {
                    return 1.524;
                    //default value 5 feet
                }
                else
                {
                    return segmentMaximumSpacing.Value;
                }
            }
        }

        /// <summary>
        ///Gets the top rail radius (meters) using the top rail section name and reference standard from the given handrail part.
        /// </summary>
        /// <param name="handrailPart">The handrail part.</param>
        /// <returns>Returns the top rail radius.</returns>
        public override double GetTopRailRadius(Part handrailPart)
        {
            //get the top rail section name and reference standard from the handrail part to calculate the radius
            string toprailSectionName = StructHelper.GetStringProperty(handrailPart, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToprailSectionName);
            string toprailSectionReferenceStandard = StructHelper.GetStringProperty(handrailPart, SPSSymbolConstants.IJUAHRTypeAProps, SPSSymbolConstants.ToprailSectionRefStandard);

            //Get the top rail crossection using the top rail section name and reference standard
            CrossSection crossSection = base.GetCrossSection(toprailSectionReferenceStandard, toprailSectionName);

            // Get toprail section depth and width from crosssection
            double topRailDepth = crossSection.Depth;
            double topRailWidth = crossSection.Width;

            // 3 Inch clearance
            double clearance = 3 * 0.0254; 

            //return the radius of the toprail using top rail width and depth
            return Math.Sqrt((topRailWidth * topRailWidth) + (topRailDepth * topRailDepth)) / 2 + clearance;

        }
        #endregion

        #region ICustomConversionToComponents
        /// <summary>
        /// Creates connections between individual components which were created in 'CreateComponents' for given bussiness object.
        /// </summary>
        /// <param name="businessObject"></param>
        public void ConnectComponents(BusinessObject businessObject)
        {
            this.handRail = (HandRail)businessObject;

            try
            {
                ReadOnlyDictionary<int, ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>>> components = base.GetComponents();
                bool curvedRailExist = CurvedRailExists(components);
                // Get All Posts
                ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>> posts = base.GetComponents((int)MemberType.Post);

                // Connect Toprails with Post
                // Get collection of toprails for each membertype index. There could be multiple toprails.
                ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>> topRailsDictionary = base.GetComponents((int)MemberType.TopRail);

                // Get collection of midrails for each membertype index. There could be multiple midrails.
                ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>> midRailsDictionary = base.GetComponents((int)MemberType.MidRail);

                // Connect toeplate with Post
                // Get collection of toeplates for each membertype index. There could be multiple toeplates.
                ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>> toePlatesDictionary = base.GetComponents((int)MemberType.ToePlate);

                // Get collection of treatments. There should be two treatments; i.e., begin and end.
                ReadOnlyDictionary<int, ReadOnlyCollection<BusinessObject>> treatmentsDictionary = base.GetComponents((int)MemberType.EndTreatment);

                // posts should be single collection of posts
                ReadOnlyCollection<BusinessObject> postCollection = posts[1];

                //----- Segments Connections ------//
                // Connect multiple toprail segments with "Axis-End" connection
                foreach (int toprailKey in topRailsDictionary.Keys)
                {
                    base.ConnectRails(postCollection, topRailsDictionary[toprailKey], curvedRailExist);
                }

                // Connect multiple midrail segments with "Axis-End" connection
                foreach (int midrailKey in midRailsDictionary.Keys)
                {
                    base.ConnectRails(postCollection, midRailsDictionary[midrailKey], curvedRailExist);
                }

                // Connect multiple toeplate segments with "Axis-End" connection
                foreach (int toePlateKey in posts.Keys)
                {
                    base.ConnectRails(postCollection, toePlatesDictionary[toePlateKey], curvedRailExist);
                }

                // Connect End Treatments
                if (treatmentsDictionary.Count > 0)
                {
                    base.ConnectEndTreatments(treatmentsDictionary, topRailsDictionary, midRailsDictionary);
                }

                //----- Rails Connections with posts ------//
                // Connect midrails with Posts
                foreach (int midrailKey in midRailsDictionary.Keys)
                {
                    base.ConnectPostsAndRails(postCollection, midRailsDictionary[midrailKey], false, curvedRailExist);
                }

                foreach (int toePlateKey in toePlatesDictionary.Keys)
                {
                    ConnectPostsAndRails(postCollection, toePlatesDictionary[toePlateKey], false, curvedRailExist);
                }

                foreach (int toprailKey in topRailsDictionary.Keys)
                {
                    ConnectPostsAndRails(postCollection, topRailsDictionary[toprailKey], true, curvedRailExist);
                }

                // Connect posts with related object
                base.ConnectPostsToStructureObjects(handRail);
            }
            catch (Exception ex)
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrConnectComponents, "Error while connecting the handrail components." + ex.ToString()));
            }
        }

        /// <summary>
        /// Creates individual components for given bussiness object.
        /// </summary>
        /// <param name="businessObject">Business object </param>
        public void CreateComponents(BusinessObject businessObject)
        {
            try
            {
                // Set the HandRail object
                this.handRail = (HandRail)businessObject;

                // Call CreateOutputs to create the components
                ConstructOutputs();
            }
            catch (Exception)
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrConstructOutputs,
                        "Error in constructing outputs for handrail in HandrailTypeA Symbol. Check custom code or contact S3D support."));
                }
            }
        }

        
        #endregion
    }

    
}

