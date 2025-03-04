using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Grids.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// PinJigCustomReport to log the data relating to the given PinJig.
    /// </summary>
    public class PinJigBasicCustomReport : ManufacturingCustomReportRuleBase
    {
        #region Private Members

        CustomReportLogService logService = null;
        int Precision = 4;
        #endregion Private Members

        /// <summary>
        /// Generates a report for the manufacturing entities to the desired path.
        /// </summary>
        /// <param name="entities">List of manufacturing entities or their parent entities for which custom report has to be generated.</param>
        /// <param name="filePath">Complete path of the output file including its name.</param>
        /// <exception cref="ArgumentNullException">Raised when the list of manufacturing entities are null.</exception>
        public override void Generate(IEnumerable<BusinessObject> entities, string filePath)
        {

            #region Validation of Inputs

            //Validates the input file path.
            ValidateFilePath(filePath);

            //Validates the input file type.
            ValidateFileType(filePath);

            //Validates the input Manufacturing Entity.
            ValidateReportingEntity(entities, ManufacturingEntityType.PinJig);

            //Check for DataBase Connection
            CheckDBConnection();

            #endregion

            foreach (BusinessObject entity in entities)
            {
                if ((entity is AssemblyBase) || (entity is PlatePartBase))
                {
                    Collection<ManufacturingBase> pinjigs = EntityService.GetManufacturingEntity(entity, ManufacturingEntityType.PinJig);
                    foreach (ManufacturingBase pinjig in pinjigs)
                    {
                        ReportPinJigObjectInformation((PinJig)pinjig, filePath);
                    }
                }
                else if (entity is PinJig)
                {
                    ReportPinJigObjectInformation((PinJig)entity, filePath);
                }
            }
        }

        /// <summary>
        /// Reports the PinJig information, given the PinJig and the file name.
        /// This method opens the file in the given filepath, calls methods for logging the information and then closes the file.
        /// </summary>
        /// <param name="pinjig">PinJig, whose information is to be generated.</param>
        /// <param name="filePath">FilePath, where the report is to be generated.</param>
        private void ReportPinJigObjectInformation(PinJig pinjig, string filePath)
        {
            logService = new CustomReportLogService();
            logService.OpenFile(filePath);

            logService.WriteToFile("PinJig Report - " + pinjig.Name);

            //Logs information related to the Angles, Girth and Straight distances between the pins, taking margin and shrinkage into consideration.
            ShowPinJigInformation(pinjig);

            //Logs the information about the RemarkingLines
            ShowRemarkingLineInformation(pinjig);

            //Logs the information related to Jig Intersection Point, taking marking and shrinkage into consideration.
            ShowJigIntxPtInformation(pinjig);

            //Logs the information related to each Pin.
            ShowMfgPinInformation(pinjig);

            //Logs the Mounting angle information.
            ShowMountingAngleInformation(pinjig);

            logService.WriteToFile("============================= End of report ===================================");
            logService.CloseFile();
        }

        #region Private Methods
        /// <summary>
        /// Logs information about PinJig, spacing between Pins, considering margins and shrinkages.
        /// </summary>
        /// <param name="pinjig">Pinjig whose information needs to be retrieved.</param>
        private void ShowPinJigInformation(PinJig pinjig)
        {
            if (pinjig != null)
            {
                PinJigReport pinjigReport = new PinJigReport(pinjig);

                logService.WriteToFile("================================================================");
                logService.WriteToFile(" ");

                logService.WriteToFile("Pin Jig Local Coordinate System Information:");

                ILocalCoordinateSystem cs = (ILocalCoordinateSystem)pinjig;
                Position csPosition = cs.Origin;

                logService.WriteToFile("Position  " + ": (" + Math.Round(csPosition.X, Precision) + ", " + Math.Round(csPosition.Y, Precision) + ", " + Math.Round(csPosition.Z, Precision) + ")");

                Vector xVector,yVector,zVector;

                xVector = cs.XAxis;
                yVector = cs.YAxis;
                zVector = cs.ZAxis;

                logService.WriteToFile("X Axis " + ": (" + Math.Round(xVector.X, Precision) + ", " + Math.Round(xVector.Y, Precision) + ", " + Math.Round(xVector.Z, Precision) + ")");
                logService.WriteToFile("Y Axis " + ": (" + Math.Round(yVector.X, Precision) + ", " + Math.Round(yVector.Y, Precision) + ", " + Math.Round(yVector.Z, Precision) + ")");
                logService.WriteToFile("Z Axis " + ": (" + Math.Round(zVector.X, Precision) + ", " + Math.Round(zVector.Y, Precision) + ", " + Math.Round(zVector.Z, Precision) + ")");

                logService.WriteToFile("================================================================");
                logService.WriteToFile(" ");

                Position pinOrigin = pinjigReport.OriginPin.GetPosition(PinPosition.JigFloor);

                logService.WriteToFile("Origin Pin Position " + ": (" + Math.Round(pinOrigin.X, Precision) + ", " + Math.Round(pinOrigin.Y, Precision) + ", " + Math.Round(pinOrigin.Z, Precision) + ")");

                logService.WriteToFile("================================================================");
                logService.WriteToFile(" ");

                logService.WriteToFile("================================================================");
                logService.WriteToFile(" ");

                double HorSpacingBetweenPins = pinjigReport.HorizontalPinInterval;
                double VerSpacingBetweenPins = pinjigReport.VerticalPinInterval;

                logService.WriteToFile("Horizontal spacing (interval) between Pins : " + Math.Round(HorSpacingBetweenPins, Precision));
                logService.WriteToFile("Vertical spacing (interval) between Pins   : " + Math.Round(VerSpacingBetweenPins, Precision));
                logService.WriteToFile(" ");

                int[] MarginTypes = new int[] { };
                PinJigDiagonalReport pinjigDiagonalReport = new PinJigDiagonalReport(pinjig, false, MarginTypes);
                double StraightDiag_OP_To_DP = pinjigDiagonalReport.OriginToDiagonalPointMinimumDistance;
                double StraightDiag_XP_To_YP = pinjigDiagonalReport.HorizontalPointToVerticalPointMinimumDistance;

                logService.WriteToFile("Straight distance between Origin Point and Diagonal Point corners : " + Math.Round(StraightDiag_OP_To_DP, Precision));
                logService.WriteToFile("Straight distance between Horizontal Point to origin and Vertical Point to origin corners : " + Math.Round(StraightDiag_XP_To_YP, Precision));
                logService.WriteToFile(" ");

                #region dump the margin & shrinkage values for all supported plates.
                logService.WriteToFile("Margin and Shrinkage inputs for the new API (obtained from Model)");

                ReadOnlyCollection<PlatePartBase> supportedPlates = pinjig.SupportedPlates;

                foreach (PlatePartBase plate in supportedPlates)
                {
                    Collection<BusinessObject> MarginPortGeom = new Collection<BusinessObject>();
                    double[] MarginValues = new double[] { };
                    Vector PrimShrDir = new Vector();
                    Vector SecShrDir = new Vector();
                    double PrimShrVal = 0, SecShrVal = 0;

                    EntityService.GetMarginData(plate, out MarginValues, out MarginPortGeom);

                    EntityService.GetShrinkageData(plate, out PrimShrVal, out PrimShrDir, out SecShrVal, out SecShrDir);

                    string plateName = plate.Name;

                    if (MarginValues.Length <= 0)
                    {
                        logService.WriteToFile("\n" + plateName + " does not have any margins defined!");
                    }
                    else
                    {

                        logService.WriteToFile("\n" + plateName + " has the following margins defined:");
                        int i = 1;
                        foreach (double marginValue in MarginValues)
                        {
                            logService.WriteToFile("(" + i + ") " + Math.Round(marginValue, Precision));
                            i++;
                        }
                    }

                    logService.WriteToFile("");

                    if (PrimShrDir != null)
                    {
                        logService.WriteToFile("//" + plateName + " has a Primary shrinkage value   = " + Math.Round(PrimShrVal, Precision) + " along (" + Math.Round(PrimShrDir.X, Precision) + ", " + Math.Round(PrimShrDir.Y, Precision) + ", " + Math.Round(PrimShrDir.Z, Precision) + ")");
                    }

                    if (SecShrDir != null)
                    {
                        logService.WriteToFile("//" + plateName + " has a Secondary shrinkage value   = " + Math.Round(PrimShrVal, Precision) + " along (" + Math.Round(SecShrDir.X, Precision) + ", " + Math.Round(SecShrDir.Y, Precision) + ", " + Math.Round(SecShrDir.Z, Precision) + ")");
                    }
                }

                Vector HorShrDir = new Vector();
                Vector VerShrDir = new Vector();

                HorShrDir.Set(1, 0, 0);
                VerShrDir.Set(0, 1, 0);

                double HorShrVal = 0, VerShrVal = 0;
                EntityService.GetShrinkageData(supportedPlates, HorShrDir, VerShrDir, out HorShrVal, out VerShrVal);

                logService.WriteToFile(" ");
                logService.WriteToFile("Amortized Shrinkage Value for all supported plates together:");
                logService.WriteToFile("Shrinkage value   = " + HorShrVal + " along Horizontal (" + Math.Round(HorShrDir.X, Precision) + ", " + Math.Round(HorShrDir.Y, Precision) + ", " + Math.Round(HorShrDir.Z, Precision) + ")");
                logService.WriteToFile("Shrinkage value = " + VerShrVal + " along Vertical (" + Math.Round(VerShrDir.X, Precision) + ", " + Math.Round(VerShrDir.Y, Precision) + ", " + Math.Round(VerShrDir.Z, Precision) + ")");

                #endregion dump the margin & shrinkage values for all supported plates.

                int[] MarginTypesAffectingCorners = new int[] { };
                double[,] Dist = new double[4, 4];

                PinJigDiagonalReport pinJigDiagonalwithoutShrReport = new PinJigDiagonalReport(pinjig, false, MarginTypesAffectingCorners);
                Dist[0, 0] = pinJigDiagonalwithoutShrReport.OriginToDiagonalPointGirthLength;
                Dist[0, 1] = pinJigDiagonalwithoutShrReport.OriginToDiagonalPointMinimumDistance;
                Dist[0, 2] = pinJigDiagonalwithoutShrReport.HorizontalPointToVerticalPointGirthLength;
                Dist[0, 3] = pinJigDiagonalwithoutShrReport.HorizontalPointToVerticalPointMinimumDistance;

                PinJigDiagonalReport pinJigDiagonalwithShrReport = new PinJigDiagonalReport(pinjig, true, MarginTypesAffectingCorners);
                Dist[1, 0] = pinJigDiagonalwithShrReport.OriginToDiagonalPointGirthLength;
                Dist[1, 1] = pinJigDiagonalwithShrReport.OriginToDiagonalPointMinimumDistance;
                Dist[1, 2] = pinJigDiagonalwithShrReport.HorizontalPointToVerticalPointGirthLength;
                Dist[1, 3] = pinJigDiagonalwithShrReport.HorizontalPointToVerticalPointMinimumDistance;

                // Get diagonal lengths considering only margin, no shrinkage

                //     HeatingMargin = 1
                //     BendingMargin = 2
                //     EndFaceMargin = 3
                //     IntercoastalMargin = 4
                //     AssyMarginSelected = 5
                //     AssyMarginDeselected = 6

                MarginTypesAffectingCorners = new int[] { 4, 2 };

                pinJigDiagonalwithoutShrReport = new PinJigDiagonalReport(pinjig, false, MarginTypesAffectingCorners);
                Dist[2, 0] = pinJigDiagonalwithoutShrReport.OriginToDiagonalPointGirthLength;
                Dist[2, 1] = pinJigDiagonalwithoutShrReport.OriginToDiagonalPointMinimumDistance;
                Dist[2, 2] = pinJigDiagonalwithoutShrReport.HorizontalPointToVerticalPointGirthLength;
                Dist[2, 3] = pinJigDiagonalwithoutShrReport.HorizontalPointToVerticalPointMinimumDistance;


                pinJigDiagonalwithShrReport = new PinJigDiagonalReport(pinjig, true, MarginTypesAffectingCorners);
                Dist[3, 0] = pinJigDiagonalwithShrReport.OriginToDiagonalPointGirthLength;
                Dist[3, 1] = pinJigDiagonalwithShrReport.OriginToDiagonalPointMinimumDistance;
                Dist[3, 2] = pinJigDiagonalwithShrReport.HorizontalPointToVerticalPointGirthLength;
                Dist[3, 3] = pinJigDiagonalwithShrReport.HorizontalPointToVerticalPointMinimumDistance;

                string[] DistStr = new string[4];
                for (int i = 0; i < 4; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        double value = Math.Round(Dist[j, i], Precision);
                        DistStr[i] = DistStr[i] + value.ToString(" 0#.000 |");
                    }
                }

                logService.WriteToFile(" ");
                logService.WriteToFile("+------------------------------------------------+--------+--------+--------+--------+");
                logService.WriteToFile("|                                                | No S/M | Only S | Only M | S & M  |");
                logService.WriteToFile("+------------------------------------------------+--------+--------+--------+--------+");
                logService.WriteToFile("| Origin Point to Diagonal Point dist   - Girth  :" + DistStr[0]);
                logService.WriteToFile("| Origin Point to Diagonal Point dist- Straight  :" + DistStr[1]);
                logService.WriteToFile("| Vertical Point to Hori Point dist     - Girth  :" + DistStr[2]);
                logService.WriteToFile("| Vertical Point to Hori Point dist  - Straight  :" + DistStr[3]);
                logService.WriteToFile("+------------------------------------------------+--------+--------+--------+--------+");
                logService.WriteToFile(" ");
                logService.WriteToFile(" ");

                double JigFloorNormalAndXaxis = pinjigReport.GetAngle(PinJigReportAngleType.BasePlaneWithGlobalYZ);
                logService.WriteToFile("Angle between Jig Floor and YZ plane : " + Math.Round(JigFloorNormalAndXaxis, Precision));

                double JigFloorNormalAndYaxis = pinjigReport.GetAngle(PinJigReportAngleType.BasePlaneWithGlobalZX);
                logService.WriteToFile("Angle between Jig Floor and ZX plane : " + Math.Round(JigFloorNormalAndYaxis, Precision));

                double JigFloorNormalAndZaxis = pinjigReport.GetAngle(PinJigReportAngleType.BasePlaneWithGlobalXY);
                logService.WriteToFile("Angle between Jig Floor and XY plane : " + Math.Round(JigFloorNormalAndZaxis, Precision));
                logService.WriteToFile(" ");
            }
        }

        /// <summary>
        /// Logs information relating to the PinJig Remarking Lines.
        /// </summary>
        /// <param name="pinjig">Pinjig whose information needs to be retrieved.</param>
        private void ShowRemarkingLineInformation(PinJig pinjig)
        {
            logService.WriteToFile("----------------------------------------------------------------");
            logService.WriteToFile(" ");

            ReadOnlyCollection<ManufacturingGeometry> remarkingLines = pinjig.GetRemarkingLines(PinJigRemarkingDataLocation.JigFloor, PinJigRemarkingLineTopologyType.All, PinJigRemarkingLineDirection.All);

            Collection<ManufacturingGeometry> remarkingLines3D = new Collection<ManufacturingGeometry>();

            //Converting 2D to 3D 
            foreach (ManufacturingGeometry remarking in remarkingLines)
            {
                remarkingLines3D.Add(pinjig.GetSurfaceRemarkingLine(remarking));
            }
            logService.WriteToFile(" ");
            logService.WriteToFile("There are " + remarkingLines.Count + " 3D remarking lines.");
            logService.WriteToFile(" ");

            int lineIdx = 1;
            foreach (ManufacturingGeometry remarkingLine in remarkingLines3D)
            {
                logService.WriteToFile("[" + lineIdx + "] " + remarkingLine.Name + " is a 3D " + PinJig.GetRemarkingLineType(remarkingLine).ToString() + " remarking line");
                lineIdx++;
            }

            ReadOnlyCollection<ManufacturingGeometry> horRemarkingLines = pinjig.GetRemarkingLines(PinJigRemarkingDataLocation.PartSurface, PinJigRemarkingLineTopologyType.All, PinJigRemarkingLineDirection.Horizontal);
            ReadOnlyCollection<ManufacturingGeometry> verRemarkingLines = pinjig.GetRemarkingLines(PinJigRemarkingDataLocation.PartSurface, PinJigRemarkingLineTopologyType.All, PinJigRemarkingLineDirection.Vertical);

            logService.WriteToFile(" ");
            logService.WriteToFile("There are " + horRemarkingLines.Count + " horizontal and " + verRemarkingLines.Count + " vertical remarking lines.");

            double horDist = 0, verDist = 0;

            foreach (ManufacturingGeometry horRemarkingLine in horRemarkingLines)
            {
                PinJigRemarkingLineReport pinjigRemarkingLineReport = new PinJigRemarkingLineReport(horRemarkingLine);
                ReadOnlyCollection<JigRemarkingIntersection> intPts = pinjigRemarkingLineReport.GetRemarkingIntersections(PinJigRemarkingIntersectionType.All);
                logService.WriteToFile(" ");
                logService.WriteToFile(horRemarkingLine.Name + " is a horizontal " + PinJig.GetRemarkingLineType(horRemarkingLine).ToString() + " remarking line with " + intPts.Count + " intersection points:");

                int lineIndex = 1;
                foreach (JigRemarkingIntersection intersectionPoint in intPts)
                {
                    PinJigPin nearestPin = intersectionPoint.NearestPin;
                    string pinName = nearestPin.Name;

                    intersectionPoint.GetJigFloorOffsets(nearestPin, out horDist, out verDist);
                    logService.WriteToFile("[ " + lineIndex + "-a ] Offset (" + Math.Round(horDist, Precision) + ", " + Math.Round(verDist, Precision) + ") from Pin " + pinName);

                    int pinRow = nearestPin.Row;
                    int pinCol = nearestPin.Column;

                    Position namedPos = pinjigRemarkingLineReport.GetPinLineIntersection(pinRow, pinCol);
                    if (namedPos != null)
                    {
                        logService.WriteToFile("[ " + lineIndex + "-b ] Pt on Remark line closest to " + pinName + ": (" + Math.Round(namedPos.X, Precision) + ", " + Math.Round(namedPos.Y, Precision) + ", " + Math.Round(namedPos.Z, Precision) + ")");
                    }

                    namedPos = pinjigRemarkingLineReport.GetPinLineIntersection(pinRow, 0);
                    if (namedPos != null)
                    {
                        logService.WriteToFile("[ " + lineIndex + "-c ] Pt on Remark line along Pin Row " + pinRow.ToString() + ": (" + Math.Round(namedPos.X, Precision) + ", " + Math.Round(namedPos.Y, Precision) + ", " + Math.Round(namedPos.Z, Precision) + ")");

                    }

                    namedPos = pinjigRemarkingLineReport.GetPinLineIntersection(0, pinCol);
                    if (namedPos != null)
                    {
                        logService.WriteToFile("[ " + lineIndex + "-d ] Pt on Remark line along Pin Col " + pinCol.ToString() + ": (" + Math.Round(namedPos.X, Precision) + ", " + Math.Round(namedPos.Y, Precision) + ", " + Math.Round(namedPos.Z, Precision) + ")");
                    }

                    lineIndex++;
                }
            }

            foreach (ManufacturingGeometry verRemarkingLine in verRemarkingLines)
            {
                PinJigRemarkingLineReport pinjigRemarkingLineReport = new PinJigRemarkingLineReport(verRemarkingLine);
                ReadOnlyCollection<JigRemarkingIntersection> intPts = pinjigRemarkingLineReport.GetRemarkingIntersections(PinJigRemarkingIntersectionType.All);
                logService.WriteToFile(" ");
                logService.WriteToFile(verRemarkingLine.Name + " is a vertical " + PinJig.GetRemarkingLineType(verRemarkingLine).ToString() + " remarking line with " + intPts.Count + " intersection points:");

                int lineIndex = 1;
                foreach (JigRemarkingIntersection intersectionPoint in intPts)
                {
                    PinJigPin nearestPin = intersectionPoint.NearestPin;
                    string pinName = nearestPin.Name;

                    intersectionPoint.GetJigFloorOffsets(nearestPin, out horDist, out verDist);
                    logService.WriteToFile("[ " + lineIndex + "-a ] Offset (" + Math.Round(horDist, Precision) + ", " + Math.Round(verDist, Precision) + ") from Pin " + pinName);

                    int pinRow = nearestPin.Row;
                    int pinCol = nearestPin.Column;

                    Position namedPos = pinjigRemarkingLineReport.GetPinLineIntersection(pinRow, pinCol);
                    if (namedPos != null)
                    {
                        logService.WriteToFile("[ " + lineIndex + "-b ] Pt on Remark line closest to " + pinName + ": (" + Math.Round(namedPos.X, Precision) + ", " + Math.Round(namedPos.Y, Precision) + ", " + Math.Round(namedPos.Z, Precision) + ")");
                    }

                    namedPos = pinjigRemarkingLineReport.GetPinLineIntersection(pinRow, 0);
                    if (namedPos != null)
                    {
                        logService.WriteToFile("[ " + lineIndex + "-c ] Pt on Remark line along Pin Row " + pinRow.ToString() + ": (" + Math.Round(namedPos.X, Precision) + ", " + Math.Round(namedPos.Y, Precision) + ", " + Math.Round(namedPos.Z, Precision) + ")");

                    }

                    namedPos = pinjigRemarkingLineReport.GetPinLineIntersection(0, pinCol);
                    if (namedPos != null)
                    {
                        logService.WriteToFile("[ " + lineIndex + "-d ] Pt on Remark line along Pin Col " + pinCol.ToString() + ": (" + Math.Round(namedPos.X, Precision) + ", " + Math.Round(namedPos.Y, Precision) + ", " + Math.Round(namedPos.Z, Precision) + ")");
                    }

                    lineIndex++;
                }
            }
        }

        /// <summary>
        /// Gets the information about Jig Intersection Point.
        /// </summary>
        /// <param name="JigRemarkingIntersection">Pinjig whose information needs to be retrieved.</param>
        private void GetJigIntxPtInformation(JigRemarkingIntersection JigRemarkingIntersection)
        {

            logService.WriteToFile("----------------------------------------------------------------");
            logService.WriteToFile(" ");

            ManufacturingGeometry horRemarkingLine = JigRemarkingIntersection.HorizontalRemarkingParent;
            if (horRemarkingLine != null)
            {
                logService.WriteToFile("    Horizontal remarking line responsible for Intersection point : " + horRemarkingLine.Name);
            }
            else
            {
                logService.WriteToFile("    Horizontal remarking line responsible for Intersection point : ");
            }

            ManufacturingGeometry verRemarkingLine = JigRemarkingIntersection.VerticalRemarkingParent;
            if (verRemarkingLine != null)
            {
                logService.WriteToFile("    Vertical remarking line responsible for Intersection point   : " + verRemarkingLine.Name);
            }
            else
            {
                logService.WriteToFile("    Vertical remarking line responsible for Intersection point   : ");
            }

            logService.WriteToFile(" ");

            Position lowerPosition = JigRemarkingIntersection.GetPosition(PinJigRemarkingDataLocation.JigFloor);

            if (lowerPosition != null)
            {
                double lowerX = lowerPosition.X;
                double lowerY = lowerPosition.Y;
                double lowerZ = lowerPosition.Z;

                logService.WriteToFile("    Intersection of above 2 lines projected onto jig floor : (" + Math.Round(lowerX, Precision) + ", " + Math.Round(lowerY, Precision) + ", " + Math.Round(lowerZ, Precision) + ")");
            }


            Position upperPosition = JigRemarkingIntersection.GetPosition(PinJigRemarkingDataLocation.PartSurface);
            if (upperPosition != null)
            {

                double upperX = upperPosition.X;
                double upperY = upperPosition.Y;
                double upperZ = upperPosition.Z;

                logService.WriteToFile("    Coordinates of intersection of above 2 remarking lines : (" + Math.Round(upperX, Precision) + ", " + Math.Round(upperY, Precision) + ", " + Math.Round(upperZ, Precision) + ")");
            }

            logService.WriteToFile(" ");

            double horGirLenNoShrink = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.HorizontalLengthFromReferenceIntersectionWithoutShrinkage);
            double horGirLenWithShrk = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.HorizontalLengthFromReferenceIntersection);
            double verGirLenNoShrink = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.VerticalLengthFromReferenceIntersectionWithoutShrinkage);
            double verGirLenWithShrk = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.VerticalLengthFromReferenceIntersection);

            logService.WriteToFile("    Horizontal Girth Length not considering Shrinkage : " + Math.Round(horGirLenNoShrink, Precision));
            logService.WriteToFile("    Horizontal Girth Length considering Shrinkage     : " + Math.Round(horGirLenWithShrk, Precision));
            logService.WriteToFile("    Vertical Girth Length not considering Shrinkage   : " + Math.Round(verGirLenNoShrink, Precision));
            logService.WriteToFile("    Vertical Girth Length considering Shrinkage       : " + Math.Round(verGirLenWithShrk, Precision));

            double angle = JigRemarkingIntersection.BendAngle;
            if (angle > 0)
            {
                logService.WriteToFile("    The cusp angle here is: " + Math.Round(angle, Precision));
            }

            if (JigRemarkingIntersection.IsJigCorner)
            {
                logService.WriteToFile("    This is a jig corner point! ");
            }

            if (JigRemarkingIntersection.IsPlateCorner)
            {
                logService.WriteToFile("    This is a Plate corner point! ");
            }


            if (JigRemarkingIntersection.IsCrossPoint)
            {
                logService.WriteToFile("    This is a remarking intersection point! ");
            }

            if (JigRemarkingIntersection.IsBendPoint)
            {
                logService.WriteToFile("    This is a cusp point! ");
            }


            if (JigRemarkingIntersection.IsJigCorner)
            {
                double horMarginValue = JigRemarkingIntersection.MarginAtHorizontalParent;
                double verMarginValue = JigRemarkingIntersection.MarginAtVerticalParent;

                logService.WriteToFile("    Margin Value at the intersection point on the Horizontal remarking line :  " + Math.Round(horMarginValue, Precision));
                logService.WriteToFile("    Margin Value at the intersection point on the Vertical remarking line   :  " + Math.Round(verMarginValue, Precision));

            }

            double horDist = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.HorizontalDistanceFromNearestPin);
            double verDist = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.VertictalDistanceFromNearestPin);
            logService.WriteToFile("    Distances from Nearest Pin:  ");
            logService.WriteToFile("      Horizontal Distance : " + Math.Round(horDist, Precision));
            logService.WriteToFile("        Vertical Distance : " + Math.Round(verDist, Precision));

            horDist = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.HorizontalDistanceFromOriginPin);
            verDist = JigRemarkingIntersection.GetLength(JigRemarkingIntersectionLengthReportType.VerticalDistanceFromOriginPin);

            logService.WriteToFile("    Distances from Origin Pin:  ");
            logService.WriteToFile("      Horizontal Distance : " + Math.Round(horDist, Precision));
            logService.WriteToFile("        Vertical Distance : " + Math.Round(verDist, Precision));

        }

        /// <summary>
        /// Logs information about Jig Intersecction Point.
        /// </summary>
        /// <param name="pinJig">Pinjig whose information needs to be retrieved.</param>
        private void ShowJigIntxPtInformation(PinJig pinJig)
        {
            ReadOnlyCollection<JigRemarkingIntersection> JigRemarkingIntersectionPoints = pinJig.GetRemarkingIntersections(PinJigRemarkingIntersectionType.All);

            int count = 1;
            JigRemarkingIntersection prevPoint = null;
            foreach (JigRemarkingIntersection JigRemarkingIntersectionPoint in JigRemarkingIntersectionPoints)
            {
                logService.WriteToFile("Information on Jig Intersection point No." + count + " of " + JigRemarkingIntersectionPoints.Count);

                GetJigIntxPtInformation(JigRemarkingIntersectionPoint);

                if (prevPoint != null)
                {
                    double xDistance = 0, yDistance = 0;
                    JigRemarkingIntersectionPoint.GetJigFloorOffsets(prevPoint, out xDistance, out yDistance);
                    logService.WriteToFile("    Offsets from previous intersection point    X: " + Math.Round(xDistance, Precision) + "  Y:" + Math.Round(yDistance, Precision));
                }

                prevPoint = JigRemarkingIntersectionPoint;
                count++;
            }

        }

        /// <summary>
        /// Gets the information of the Pin
        /// </summary>
        /// <param name="pin">Pinjig whose information needs to be retrieved.</param>
        private void GetMfgPinInformation(PinJigPin pin)
        {
            logService.WriteToFile("----------------------------------------------------------------");
            logService.WriteToFile(" ");

            double lowerX = 0, lowerY = 0, lowerZ = 0;
            double upperX = 0, upperY = 0, upperZ = 0;

            Position pinPosition = pin.GetPosition(PinPosition.JigFloor);
            lowerX = pinPosition.X;
            lowerY = pinPosition.Y;
            lowerZ = pinPosition.Z;

            logService.WriteToFile("    Base coordinates of Pin (on jig floor)                        : (" + Math.Round(lowerX, Precision) + ", " + Math.Round(lowerY, Precision) + ", " + Math.Round(lowerZ, Precision) + ")");

            pinPosition = pin.GetPosition(PinPosition.SupportedSurface);
            upperX = pinPosition.X;
            upperY = pinPosition.Y;
            upperZ = pinPosition.Z;

            logService.WriteToFile("    Coordinates of Pin (Supported Point)                          : (" + Math.Round(upperX, Precision) + ", " + Math.Round(upperY, Precision) + ", " + Math.Round(upperZ, Precision) + ")");

            logService.WriteToFile("    Distance between upper and lower point (should be = height)   : " + Math.Round(Math.Sqrt((lowerX - upperX) * (lowerX - upperX) + (lowerY - upperY) * (lowerY - upperY) + (lowerZ - upperZ) * (lowerZ - upperZ)), Precision));

            double pinHeight = pin.GetDistance(PinDistanceReportType.Height);
            logService.WriteToFile("    Height of pin                                                 : " + Math.Round(pinHeight, Precision));
            logService.WriteToFile(" ");

            if (pinHeight > 0)
            {
                double pinHeadHorAngle = pin.HorizontalAngleAtTip;
                logService.WriteToFile("    Angle between Surface-Tangent at contact-point and horizontal : " + Math.Round(pinHeadHorAngle, Precision));

                double pinHeadVerAngle = pin.VerticalAngleAtTip;
                logService.WriteToFile("    Angle between Surface-Tangent at contact-point and vertical   : " + Math.Round(pinHeadVerAngle, Precision));
                logService.WriteToFile(" ");

                double seamDistance = pin.GetDistance(PinDistanceReportType.NearestLeftOrRightBoundary);
                logService.WriteToFile("    Distance of Pin from nearest seam                             : " + Math.Round(seamDistance, Precision));
                logService.WriteToFile(" ");

                double buttDistance = pin.GetDistance(PinDistanceReportType.NearestBottomOrTopBoundary);
                logService.WriteToFile("    Distance of Pin from nearest butt                             : " + Math.Round(buttDistance, Precision));
                logService.WriteToFile(" ");
            }

            logService.WriteToFile(" ");
        }

        /// <summary>
        /// Logs the information about the Pin.
        /// </summary>
        /// <param name="pinjig">Pinjig whose information needs to be retrieved.</param>
        private void ShowMfgPinInformation(PinJig pinjig)
        {
            PinJigPin[,] rowsOfPins = pinjig.Pins;
            logService.WriteToFile("There are " + rowsOfPins.GetLength(0) + " rows of Pins");

            int rowCount = rowsOfPins.GetLength(0);
            int coloumnCount = rowsOfPins.GetLength(1);
            int[] rowArr = new int[5];

            if (rowCount > 4)
            {
                double midRow = (double) rowCount / 2;
                rowArr[0] = 0;
                rowArr[1] = 1;
                rowArr[2] = (int)Math.Round(midRow);
                rowArr[3] = rowCount - 2;
                rowArr[4] = rowCount - 1;
            }
            else
            {
                for (int i = 0; i < rowCount; i++)
                {
                    rowArr[i] = i;
                }
            }

            foreach (int row in rowArr)
            {
                PinJigPin[] pinsInRow = new PinJigPin[coloumnCount];
                for (int col = 0; col < coloumnCount; col++)
                {
                    pinsInRow[col] = rowsOfPins[row, col];
                }

                PinJigPin prevPin = null;
                int colNumber = 1;
                foreach (PinJigPin pin in pinsInRow)
                {
                    logService.WriteToFile("Information on Pin No." + colNumber + " of " + coloumnCount + " Pins in Row No." + (row + 1));
                    GetMfgPinInformation(pin);
                    if (prevPin != null)
                    {
                        double xDistance = 0, yDistance = 0;
                        pin.GetJigFloorOffsets(prevPin, out xDistance, out yDistance);
                        logService.WriteToFile("    Offsets from previous pin    X: " + Math.Round(xDistance, Precision) + "  Y:" + Math.Round(yDistance, Precision));
                    }

                    prevPin = pin;
                    colNumber++;
                }
            }

            int coloumnNumber = -1;

            if ((coloumnCount % 2) == 0)
            {
                coloumnNumber = (coloumnCount / 2) - 1;
            }
            else
            {
                coloumnNumber = (coloumnCount / 2);
            }


            PinJigPin prevousPin = null;
            for (int i = 0; i < rowCount; i++)
            {
                PinJigPin rowPin = rowsOfPins[i, coloumnNumber];
                logService.WriteToFile("Information on Pin No." + (coloumnNumber + 1) + " of " + coloumnCount + " Pins in Row No." + (i + 1));
                GetMfgPinInformation(rowPin);

                if (prevousPin != null)
                {
                    double xDistance = 0, yDistance = 0;
                    rowPin.GetJigFloorOffsets(prevousPin, out xDistance, out yDistance);
                    logService.WriteToFile("    Offsets from previous pin    X: " + Math.Round(xDistance, Precision) + "  Y:" + Math.Round(yDistance, Precision));
                }
                prevousPin = rowPin;
            }
        }

        /// <summary>
        /// Logs the MountingAngleInformation.
        /// </summary>
        private void ShowMountingAngleInformation(PinJig pinjig)
        {
            

            PinJigOrientationDefinition orientationDefinition = pinjig.OrientationDefinition;
            IPlane basePlane = orientationDefinition.BasePlane;

            Vector basePlaneNormal = new Vector(basePlane.Normal);           
            basePlaneNormal.Length = 1;

            Vector xVector = new Vector(1, 0, 0);
            Vector yVector = new Vector(0, 1, 0);
            Vector zVector = new Vector(0, 0, 1);

            double dotX, dotY, dotZ;

            dotX = Math.Abs(xVector.Dot(basePlaneNormal));
            dotY = Math.Abs(yVector.Dot(basePlaneNormal));
            dotZ = Math.Abs(zVector.Dot(basePlaneNormal));

            Vector parallelVector, perpendicularVector1, perpendicularVector2;

            if((dotX > dotY) && (dotX > dotZ))
            {
                parallelVector = xVector;
                perpendicularVector1 = yVector;
                perpendicularVector2 = zVector;
            }
            else if ((dotY > dotX) && (dotY > dotZ))
            {
                parallelVector = yVector;
                perpendicularVector1 = xVector;
                perpendicularVector2 = zVector;
            }
            else
            {
                parallelVector = zVector;
                perpendicularVector1 = xVector;
                perpendicularVector2 = yVector;
            }

            ReadOnlyCollection<GridPlaneBase> framesCollection1 = EntityService.GetPlanesInRange(pinjig.RemarkingSurface, perpendicularVector1, pinjig.CoordinateSystemFromParent, false);
            if (framesCollection1 == null)
            { 
                return; 
            }

            ReadOnlyCollection<GridPlaneBase> framesCollection2 = EntityService.GetPlanesInRange(pinjig.RemarkingSurface, perpendicularVector2, pinjig.CoordinateSystemFromParent, false);
            if (framesCollection2 == null || framesCollection2.Count == 0)
            { 
                return; 
            }
           
            logService.WriteToFile(" ");
            logService.WriteToFile("Reporting Mounting Angle of the frame relative to PinJig base plane normal at given position: ");

            logService.WriteToFile("Intersecting frame to get the mounting position : " + framesCollection2[0].ToString());   //framesCollection2[0] is the intersector
            logService.WriteToFile("----------------------------------------------------------------");
            logService.WriteToFile(" ");

            double concaveRadius = 0.2;
            double convexRadius = 0.1;

            for (int i = 0; i<framesCollection1.Count; i++)
            {
                PinJigReport pinJigReport = pinjig.Report;
                Position mountingPosition;
                double mountingAngle1, mountingAngle2;
                const double pi = 3.14159265358979;

                double mountingAngle = pinJigReport.GetMountingAngle(framesCollection1[i], framesCollection2[0], out mountingPosition);

                logService.WriteToFile("Mounting Angle of " + framesCollection1[i] + " : " + Math.Round(mountingAngle * (180 / pi), Precision));

                pinJigReport.GetMountingAngle(framesCollection1[i], framesCollection2[0], convexRadius, concaveRadius, out mountingPosition, out mountingAngle1, out mountingAngle2);

                logService.WriteToFile("Mounting Angles of " + framesCollection1[i] + " based on  - " + "Concave radius : " + concaveRadius + " ; Convex radius : " + convexRadius);
                logService.WriteToFile("Angle1 :" + Math.Round(mountingAngle1 * (180 / pi), Precision) + " ; Angle2 :" + Math.Round(mountingAngle2 * (180 / pi), Precision));
                logService.WriteToFile("Mounting Position : (" + Math.Round(mountingPosition.X, Precision) + ", " + Math.Round(mountingPosition.Y, Precision) + ", " + Math.Round(mountingPosition.Z, Precision) + ")");
                logService.WriteToFile(" ");
            }            

            logService.WriteToFile("================================================================");

            ShowMountingAngleInformationBasedOnInputs(pinjig, PinJigRemarkingLineTopologyType.Plate);

            ShowMountingAngleInformationBasedOnInputs(pinjig, PinJigRemarkingLineTopologyType.Profile);
        }



        /// <summary>
        /// This method is to log the mounting angle of the input part(Plate or Profile)
        /// </summary>
        /// <param name="pinjig">Input pinjig.</param>
        /// <param name="remarkingLineTopologyType">Type of remarkingLineTopology</param>
        private void ShowMountingAngleInformationBasedOnInputs(PinJig pinjig, PinJigRemarkingLineTopologyType remarkingLineTopologyType)
        {
            // Validate inputs
            if (pinjig == null) return;

            // If PinJigRemarkingLineTopologyType is not plate or profile then return.
            string partType = string.Empty;
            if (remarkingLineTopologyType == PinJigRemarkingLineTopologyType.Plate)
            {
                partType = "plate part";
            }
            else if (remarkingLineTopologyType == PinJigRemarkingLineTopologyType.Profile)
            {
                partType = "profile part";
            }
            else
            {
                return;
            }

            // get remarking lines based on PinJigRemarkingLineTopologyType. If no remarking line then return.     
            ReadOnlyCollection<ManufacturingGeometry> remarkingLines = pinjig.GetRemarkingLines(PinJigRemarkingDataLocation.PartSurface, remarkingLineTopologyType, PinJigRemarkingLineDirection.All);
            if (remarkingLines.Count == 0)
            {
                return;
            }

            logService.WriteToFile(" ");
            logService.WriteToFile("Reporting Mounting Angle of the " + partType + " relative to PinJig base plane normal at mid position of remarking line: ");
            logService.WriteToFile("--------------------------------------------------------------------------------------------------------------------------");
            PinJigReport pinjigReport = new PinJigReport(pinjig);
            foreach (ManufacturingGeometry remarkingLine in remarkingLines)
            {
                try
                {
                    BusinessObject part = null;
                    BusinessObject representedEntity = remarkingLine.RepresentedEntity;
                    if (representedEntity is PlateSystemBase)
                    {
                        PlateSystemBase plateSystem = (PlateSystemBase)representedEntity;
                        part = plateSystem.GetPlateParts()[0];
                    }
                    else if (representedEntity is ProfileSystem)
                    {
                        ProfileSystem profileSystem = (ProfileSystem)representedEntity;
                        part = profileSystem.GetParts()[0];
                    }
                    else
                    {
                        part = representedEntity;
                    }

                    Position centroid = remarkingLine.Geometry.Centroid;
                    Position mountingPosition = new Position(0, 0, 0);
                    double mountingAngle = pinjigReport.GetMountingAngle(part, centroid, out mountingPosition);
                    logService.WriteToFile(" ");
                    logService.WriteToFile("Mounting Angle of " + part.ToString() + " : " + Math.Round(mountingAngle * (180 / Math.PI), 2));

                    double mountingAngle1 = 0;
                    double concaveRadius = 0.2;
                    double convexRadius = 0.1;
                    pinjigReport.GetMountingAngle(part, centroid, convexRadius, concaveRadius, out mountingPosition, out mountingAngle, out mountingAngle1);
                    logService.WriteToFile("Mounting Angles of " + part.ToString() + " based on  - " + "Concave radius : " + concaveRadius + " ; Convex radius : " + convexRadius);
                    logService.WriteToFile("Angle1 :" + Math.Round(mountingAngle * (180 / Math.PI), 2) + " ; Angle2 :" + Math.Round(mountingAngle1 * (180 / Math.PI), 2));
                    logService.WriteToFile("Mounting Position : (" + Math.Round(mountingPosition.X, 2) + ", " + Math.Round(mountingPosition.Y, 2) + ", " + Math.Round(mountingPosition.Z, 2) + ")");
                }
                catch
                {
                    continue;
                }
            }
        }

        #endregion Private Methods
    }
}