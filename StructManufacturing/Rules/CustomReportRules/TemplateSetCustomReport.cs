using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Content.Manufacturing.Services;
using System.IO;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Manufacturing.Exceptions;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// TemplateSetCustomReport to log the data relating to the given TemplateSet.
    /// </summary>
    public class TemplateSetCustomReport : ManufacturingCustomReportRuleBase
    {
        #region Private Members

        CustomReportLogService logService = null;        
        private const int Precision = 4;

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
            ValidateReportingEntity(entities, ManufacturingEntityType.TemplateSet);

            //Check for DataBase Connection
            CheckDBConnection();

            #endregion

            foreach (BusinessObject entity in entities)
            {
                if (entity is IManufacturable)
                {
                    Collection<ManufacturingBase> templateSets = EntityService.GetManufacturingEntity(entity, ManufacturingEntityType.TemplateSet);
                    foreach (ManufacturingBase templateSet in templateSets)
                    {                        
                        ReportTemplateSetInformation((TemplateSet)templateSet, filePath);
                    }
                }
                else if (entity is TemplateSet)
                {
                    ReportTemplateSetInformation((TemplateSet)entity, filePath);
                }
            }
        }
        
        private void ReportTemplateSetInformation(TemplateSet templateSet, string filePath)
        {            

            logService = new CustomReportLogService();
            logService.OpenFile(filePath);

            logService.WriteToFile("TemplateSet Report - " + templateSet.Name);

            logService.WriteToFile("Number of Template Groups - " + templateSet.GroupCount);


            for (int groupNumber = 1; groupNumber <= templateSet.GroupCount; groupNumber++)
            {
                ReadOnlyCollection<Template> templates = templateSet.GetTemplates(groupNumber);

                if (templates == null)
                {
                    continue;
                }               

                logService.WriteToFile("Number of Templates in Group" + groupNumber + " : " + templates.Count);

                int count = 1;
                foreach (Template template in templates)
                {

                    logService.WriteToFile("");
                    logService.WriteToFile("Template " + template.Name + count);

                    TemplateReport templateRep = new TemplateReport(template);

                    //Logs the information regarding the position of each Template.
                    ShowTemplatePositionData(templateRep);

                    //Logs data about the Template Distance Type data for each template.
                    ShowTemplateDistanceTypeData(templateRep);

                    //"GetReferenceCurveData" needs to be Implemented
                    ShowTemplateReferenceCurveData(templateRep, templateSet);

                    //Logs the Template Angle data.
                    ShowTemplateAngleData(templateRep);

                    //Logs the Seam Entity information of the Template.
                    ShowTemplateBoundaryData(templateRep);

                    //Logs template off set data
                    ShowTemplateOffSetData(templateRep);

                    //Logs the Template height information.
                    ShowTemplateHeightData(templateRep);

                    //Logs the Template chord height information.
                    ShowChordHeightsData(template);

                    //Logs template marks data
                    ShowTemplateMarksData(templateRep);

                    //Logs the template drawing related information
                    ShowTemplateDrawingInformation(templateRep);

                    //Logs the templates intersection information
                    ShowTemplateIntersectionInformation(template);

                    count++;
                }
            }

            //Logs the TemplateSet height information.
            ShowChordHeightsData(templateSet);

            logService.WriteToFile("============================= End of report ===================================");
            logService.CloseFile();
        }


        #region Private Methods

        /// <summary>
        /// Logs the information regarding the position of each Template.
        /// </summary>
        /// <param name="templateRep">TemplateReport</param>
        private void ShowTemplatePositionData(TemplateReport templateRep)
        {

            logService.WriteToFile(" ");
            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - POINTS ");
            logService.WriteToFile("================================================================");

            Position pos = null;
            string str = "";

            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.AftSeam);
            }
            catch(Exception ce)
            {
                str = ce.ToString();
            }

            if( pos != null )            
             logService.WriteToFile("AftSeamPoint                                 : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));

            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.ForeSeam);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if( pos != null )           
                logService.WriteToFile("ForeSeamPoint                                : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));

            
            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.UpperEnd);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("UpperEndMarkingLinePoint                     : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.UpperEndOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("UpperEndTopLinePoint                         : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.UpperSeam);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("UpperSeamPoint                               : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));

            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.UpperSeamOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if( pos != null )            
                logService.WriteToFile("UpperSeamTopLinePoint                        : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.BaseControl);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("BaseControlPoint                             : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.BaseControlOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("BaseControlTopLinePoint                      : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.LowerSeam);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if( pos != null )           
                logService.WriteToFile("LowerSeamPoint                               : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));

            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.LowerSeamOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("LowerSeamTopLinePoint                        : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.LowerEnd);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("LowerEndMarkingLinePoint                     : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));

            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.LowerEndOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("LowerEndTopLinePoint                         : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.BaseControlOnChordLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("BaseControlChordPoint                        : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));

            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.LowerSeamOnChordLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("LowerSeamChordPoint                          : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.UpperSeamOnChordLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("UpperSeamChordPoint                          : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.LowerEndOnChordLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("LowerEndChordPoint                           : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.UpperEndOnChordLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("UpperEndChordPoint                           : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.AftEndOnControlLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (pos != null)
                logService.WriteToFile("AftEndOnControlLine                          : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));


            try
            {
                pos = templateRep.GetPosition(TemplatePositionType.ForeEndOnControlLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if( pos != null )            
                logService.WriteToFile("ForeEndOnControlLine                         : " + RoundedToPrecision(pos.X, Precision) + ", " + RoundedToPrecision(pos.Y, Precision) + ", " + RoundedToPrecision(pos.Z, Precision));
           

        }
        
        /// <summary>
        /// Logs data about the Template Distance Type data for each template.
        /// </summary>
        /// <param name="templateRep">TemplateReport</param>
        private void ShowTemplateDistanceTypeData(TemplateReport templateRep)
        {
            string str = "";
            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - DISTANCE ");
            logService.WriteToFile("================================================================");

            try
            {
                logService.WriteToFile("GirthAtLowerSeam                             : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.GirthBetweenLowerSeamAndReferenceLowerSeam), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

           try
            {
                logService.WriteToFile("GirthAtUpperSeam                             : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.GirthBetweenUpperSeamAndReferenceUpperSeam), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }            
            

            try
            {
                logService.WriteToFile("GirthAtBaseControlLine                       : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.GirthBetweenBaseControlAndReferenceBaseControl), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }
            

            try
            {
                logService.WriteToFile("GirthBetweenUpperSeamAndBaseCtlLine          : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamAndBaseControlAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }            

            try
            {
                logService.WriteToFile("GirthBetweenBaseCtlPointAndAftPoint          : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.BaseControlAndAftSeamAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }
            

            try
            {
                logService.WriteToFile("GirthBetweenBaseCtlPointAndForePoint         : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.BaseControlAndForeSeamAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }
            
            try
            {
                logService.WriteToFile("GirthBetweenBaseCtlPointAndCtlLineAftEndPt   : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.BaseControlAndControlLineAftEndAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("GirthBetweenBaseCtlPointAndCtlLineForeEndPt  : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.BaseControlAndControlLineForeEndAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            

            try
            {
                logService.WriteToFile("BetweenLowerSeamAndBaseCtlLine               : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamAndBaseControl), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenUpperSeamAndBaseCtlLine               : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamAndBaseControl), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerSeamAndClosestTemplate           : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamAndClosestSectionLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenUpperSeamAndClosestTemplate           : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamAndClosestSectionLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerSeamAndBaseCtlLineOnTopLine      : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamOnTopLineAndBaseControlOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenBaseCtlLineAndUpperSeamOnTopLine      : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.BaseControlOnTopLineAndUpperSeamOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerSeamAndUpperSeamOnTopLine        : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamOnTopLineAndUpperSeamOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerEndAndLowerSeamOnTopLine         : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LoweEndOnTopLineAndLowerSeamOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenUpperSeamAndUpperEndOnTopLine         : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamOnTopLineAndUpperEndOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerEndAndUpperEndOnTopLine          : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerEndOnTopLineAndUpperEndOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenSeamAndTemplate                       : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.AftSeamOnTopLineAndBaseControlOnTopLine), Precision));

            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }          

            logService.WriteToFile("(This is linear distance between Aft butt - BCL intersection and Template BCL bottom point, measured/projected along base plane)" + "\r\n");

            try
            {
                logService.WriteToFile("FromStartingEdgeToBaseCtlLine                : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerEndOnTopLineAndBaseControlOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("HU_HeightFromGroundFloor                     : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.HeightAtUpperSeamFromFloor), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("HL_HeightFromGroundFloor                     : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.HeightAtLowerSeamFromFloor), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("MaxHeightFromButtLine ( F )                  : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.MaximumHeightOfSeamChordLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }
                     
            logService.WriteToFile(" ");
            logService.WriteToFile("New enums ");

            try
            {
                logService.WriteToFile("BetweenLowerEndAndBaseCtlLine                : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerEndAndBaseControl), Precision));

            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            try
            {
                logService.WriteToFile("BetweenUpperEndAndBaseCtlLine                : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperEndAndBaseControl), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("GirthAtLowerSeam                             : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.GirthBetweenLowerSeamAndReferenceLowerSeam), Precision)); logService.WriteToFile("BetweenLowerEndAndLowerSeam                  : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerEndAndLowerSeam), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenUpperSeamAndUpperEnd                  : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamAndUpperEnd), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerEndAndUpperEnd                   : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerEndAndUpperEnd), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerSeamAndUpperSeam                 : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamAndUpperSeam), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenLowerSeamAndBaseCtlLine               : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamAndBaseControl), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("BetweenUpperSeamAndBaseCtlLine               : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamAndBaseControl), Precision));

            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            try
            {
                logService.WriteToFile("GirthBetweenLowerEndAndBaseCtlLine           : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerEndAndBaseControlAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("GirthBetweenUpperEndAndBaseCtlLine           : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperEndAndBaseControlAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }
                        
                    
            try
            {
                logService.WriteToFile("GirthAtLowerSeam                             : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.GirthBetweenLowerSeamAndReferenceLowerSeam), Precision)); logService.WriteToFile("GirthBetweenUpperSeamAndUpperEnd             : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamAndUpperEndAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("GirthBetweenLowerEndAndUpperEnd              : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerEndAndUpperEndAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("GirthBetweenLowerSeamAndUpperSeam            : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamAndUpperSeamAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("GirthBetweenLowerSeamAndBaseCtlLine          : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.LowerSeamAndBaseControlAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            
            try
            {
                logService.WriteToFile("GirthAtLowerSeam                             : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.GirthBetweenLowerSeamAndReferenceLowerSeam), Precision)); logService.WriteToFile("GirthBetweenUpperSeamAndBaseCtlLine          : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.UpperSeamAndBaseControlAsGirth), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateRep">TemplateReport</param>
        /// <param name="templateSet">TemplateSet</param>
        private void ShowTemplateReferenceCurveData(TemplateReport templateRep, TemplateSet templateSet)
        {

            ReadOnlyCollection<BusinessObject> referenceCurves = null;
            string str = "";
            try
            {
                referenceCurves = templateRep.GetReferenceCurves(TemplateReferenceCurveType.All);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (referenceCurves != null)
            {
                logService.WriteToFile("================================================================");
                logService.WriteToFile(" TEMPLATE REPORT - GET REFERENCE ");
                logService.WriteToFile("================================================================");
                logService.WriteToFile("Number of Reference Curves               : " + referenceCurves.Count);
                logService.WriteToFile(" ");

                string strObjectName = " ";
                string strBaseName = " ";
                int refCount = referenceCurves.Count;
                int templateCount = 1;
                int knuckleCount = 1;

                foreach (BusinessObject refCurve in referenceCurves)
                {
                    NamedItemHelper refCurveName = new NamedItemHelper(refCurve);
                    string strBaseControlLine = "TemplateControlLine";

                    if (refCurveName != null)
                    {
                        strObjectName = refCurveName.Name;
                    }

                    if (refCurve.SupportsInterface("IJRefCurveData"))
                    {
                        strBaseName = "Reference Curve " + ":";
                        strObjectName = refCurve.ToString();
                    }
                    else if (refCurve.SupportsInterface("IJDModelBody"))
                    {
                        strBaseName = "Knuckle Line " + knuckleCount + ":";
                        knuckleCount = knuckleCount + 1;
                    }
                    else if (refCurve.SupportsInterface("IJMfgMarkingLines_AE"))
                    {
                        strBaseName = "Marking Line:";
                    }
                    else if (refCurve.SupportsInterface("IJMfgGeom3d"))
                    {
                        strBaseName = "";
                        if (!strObjectName.Equals(strBaseControlLine))
                        {
                            strObjectName = strObjectName + " " + templateCount + ":";
                            templateCount = templateCount + 1;
                        }
                        else
                        {
                            strObjectName = strBaseControlLine + ":";
                        }
                    }

                    logService.WriteToFile(strBaseName + strObjectName);          
                    
                    //TO BE IMPLEMENTED

                    //uncomment below code if  public void GetReferenceCurveData(BusinessObject referenceCurve, out double girthLength, out double height,
                    //                                    out double projectedLength, out Position intersectionPosition) method implemented

                    /* double girthLength,projectedLength,height;
                    Position intersectionPosition;
                    templateRep.GetReferenceCurveData(refCurve, out girthLength, out height, out projectedLength, out intersectionPosition );

                    logService.WriteToFile("Reference Curve Girh Length               : " + ManufacturingATPUtils.ReallyRound(girthLength, Precision));
                    logService.WriteToFile("Reference Curve Straight Length           : " + ManufacturingATPUtils.ReallyRound(height, Precision));
                    logService.WriteToFile("Reference Curve Height                    : " + ManufacturingATPUtils.ReallyRound(projectedLength, Precision));
                    logService.WriteToFile("Reference Curve Position                  : " + "(" + ManufacturingATPUtils.ReallyRound(intersectionPosition.X, Precision) + ", " + ManufacturingATPUtils.ReallyRound(intersectionPosition.Y, Precision) + ", " + ManufacturingATPUtils.ReallyRound(intersectionPosition.Z, Precision) + ")"); */
                }

            }

        }

        /// <summary>
        /// Logs the Template Angle data.
        /// </summary>
        /// <param name="templateRep">The template rep.</param>
        private void ShowTemplateAngleData(TemplateReport templateRep)
        {

            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - GET ANGLE ");
            logService.WriteToFile("================================================================");

            TemplateSide mfgTemplateSide = TemplateSide.Aft;
            TemplateProjectionSide mfgProjectSide = TemplateProjectionSide.PositiveX;
            double attachedangle = 0.0;
            double radius = 0.2;
            string str = "";

            try
            {
                attachedangle = templateRep.GetAngle(TemplateAngleType.Default, radius, out mfgTemplateSide, out mfgProjectSide);
            }
            catch(Exception ce)
            {
                str = ce.ToString();
            }
            
            logService.WriteToFile("AttachedAngle : " + RoundedToPrecision(attachedangle, Precision));
            logService.WriteToFile(" ");

            if (mfgTemplateSide == TemplateSide.Aft)
            {
                logService.WriteToFile("Template Side : Aft ");
            }
            else if (mfgTemplateSide == TemplateSide.Fore)
            {
                logService.WriteToFile("Template Side : Fore ");
            }
            else if (mfgTemplateSide == TemplateSide.Lower)
            {
                logService.WriteToFile("Template Side : Lower ");
            }
            else if (mfgTemplateSide == TemplateSide.Upper)
            {
                logService.WriteToFile("Template Side : Upper ");
            }

            if (mfgProjectSide == TemplateProjectionSide.NegativeX)
            {
                logService.WriteToFile("Project Side  : Fore");
            }
            else if (mfgProjectSide == TemplateProjectionSide.PositiveX)
            {
                logService.WriteToFile("Project Side  : Aft");
            }
            else if (mfgProjectSide == TemplateProjectionSide.NegativeY)
            {
                logService.WriteToFile("Project Side  : StarBoard");
            }
            else if (mfgProjectSide == TemplateProjectionSide.PositiveY)
            {
                logService.WriteToFile("Project Side  : Port");
            }
            else if (mfgProjectSide == TemplateProjectionSide.NegativeZ)
            {
                logService.WriteToFile("Project Side  : Down");
            }
            else if (mfgProjectSide == TemplateProjectionSide.PositiveZ)
            {
                logService.WriteToFile("Project Side  : Up");
            }                        

        }

        /// <summary>
        /// Logs the Seam Entity information of the Template.
        /// </summary>
        /// <param name="templateRep">The template rep.</param>
        private void ShowTemplateBoundaryData(TemplateReport templateRep)
        {
           
            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - GETSEAM");
            logService.WriteToFile("================================================================");

            BusinessObject boundary = null;
            string str = "";
            
            try
            {
                boundary = templateRep.GetBoundary(TemplateSide.Aft);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            string strSeamName = BoundaryName(boundary);
            logService.WriteToFile("TemplateAftSeamName                      : " + strSeamName);

            try
            {
                boundary = templateRep.GetBoundary(TemplateSide.Fore);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            strSeamName = BoundaryName(boundary);
            logService.WriteToFile("TemplateForwardSeamName                  : " + strSeamName);

            try
            {
                boundary = templateRep.GetBoundary(TemplateSide.Upper);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            strSeamName = BoundaryName(boundary);
            logService.WriteToFile("TemplateUpperSeamName                    : " + strSeamName);

            try
            {
                boundary = templateRep.GetBoundary(TemplateSide.Lower);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            strSeamName = BoundaryName(boundary);
            logService.WriteToFile("TemplateLowerSeamName                    : " + strSeamName);
            logService.WriteToFile(" ");

        }

        /// <summary>
        /// Logs the Template height information.
        /// </summary>
        /// <param name="templateRep">The template rep.</param>
        private void ShowTemplateHeightData(TemplateReport templateRep)
        {

            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - GET OFFSET HEIGHT ");
            logService.WriteToFile("================================================================");

            double mInterval = 0.0, dInterval = 0.0;
            double height = 0.0;
            string str = "";
            try
            {
                mInterval = templateRep.GetDistance(TemplateDistanceType.LowerEndOnTopLineAndUpperEndOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            while (dInterval < mInterval)
            {
                try
                {
                    height = templateRep.GetHeight(dInterval);
                    height = RoundedToPrecision(height, Precision);
                }
                catch( Exception ce)
                {
                    str = ce.ToString();
                }               

                logService.WriteToFile("Interval   : " + dInterval.ToString("0.0") + " Value            : " + RoundedToPrecision(height, Precision));
                dInterval = dInterval + 0.2;
            }

            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - GET OFFSET HEIGHT LIST");
            logService.WriteToFile("================================================================");

            double[] offset;           
            long nSteps;

            nSteps = (int)(mInterval/0.2);
            offset = new double[nSteps+1];

            for (int count = 0; count < nSteps+1; count++)
            {
                offset[count] = count * 0.2;
                try
                {
                    height = templateRep.GetHeight(count * 0.2);
                }
                catch(Exception ce)
                {
                    str = ce.ToString();
                }               

                logService.WriteToFile("Interval   : " + RoundedToPrecision(offset[count], Precision).ToString("0.0") + " Value            : " + RoundedToPrecision(height, Precision));
            }
        }

        /// <summary>
        /// Logs the Template chord height information.
        /// </summary>
        /// <param name="reportObj">TemplateSet or Template</param>
        private void ShowChordHeightsData(BusinessObject reportObj)
        {

            ReadOnlyCollection<TemplateChordHeightData> chordHeights = null;
            string str = "";

            if (reportObj is TemplateSet)
            {
                logService.WriteToFile(" ");
                logService.WriteToFile("================================================================");
                logService.WriteToFile(" TEMPLATE SET REPORT - GET CHORD HEIGHT ");
                logService.WriteToFile("================================================================");

                TemplateSetReport templateSetRep = new TemplateSetReport((TemplateSet)reportObj);
                try
                {
                    chordHeights = templateSetRep.GetBaseControlLineHeights(1);
                }
                catch (Exception ce)
                {
                    str = ce.ToString();
                }  
            }
            else if (reportObj is Template)
            {
                logService.WriteToFile(" ");
                logService.WriteToFile("================================================================");
                logService.WriteToFile(" TEMPLATE REPORT - GET CHORD HEIGHT ");
                logService.WriteToFile("================================================================");

                TemplateReport templateSetRep = new TemplateReport((Template)reportObj);
                try
                {
                    chordHeights = templateSetRep.GetBottomLineHeights(1);
                }
                catch (Exception ce)
                {
                    str = ce.ToString();
                }
            }
            else
            {
                throw new ArgumentException(" Invalid Argument Passed to get report ");
            }

            logService.WriteToFile(" ");
            logService.WriteToFile("Length(Str):\tLength(Curve):\tHeight:\t\tPoint(Str)\t\t\tPoint(Curve)");
            
            double accumulatedOffset, girthLength, chordHeight;
            Position posOnStraightLine, posOnCurve;
            string str1, str2, str3, str4, str5, str6, str7, str8, str9;

            if (chordHeights == null)
            {
                throw new Exception(" Exception : chordHeights is empty ");
            }

            foreach (TemplateChordHeightData templateChordHeight in chordHeights)
            {
                accumulatedOffset = templateChordHeight.AccumulatedOffset;
                girthLength = templateChordHeight.GirthLength;
                chordHeight = templateChordHeight.Height;
                posOnStraightLine = templateChordHeight.PositionOnTopLine;
                posOnCurve = templateChordHeight.PositionOnChord;

                str1 = RoundedToPrecision(accumulatedOffset, Precision).ToString();
                str2 = RoundedToPrecision(girthLength, Precision).ToString();
                str3 = RoundedToPrecision(chordHeight, Precision).ToString();
                str4 = RoundedToPrecision(posOnStraightLine.X, Precision).ToString();
                str5 = RoundedToPrecision(posOnStraightLine.Y, Precision).ToString();
                str6 = RoundedToPrecision(posOnStraightLine.Z, Precision).ToString();
                str7 = RoundedToPrecision(posOnCurve.X, Precision).ToString();
                str8 = RoundedToPrecision(posOnCurve.Y, Precision).ToString();
                str9 = RoundedToPrecision(posOnCurve.Z, Precision).ToString();

                logService.WriteToFile(str1 + "\t" + "\t" + str2 + "\t" + "\t" + str3 + "\t" + "\t" + str4 + ", " + str5 + ", " + str6 + "\t" + "\t" + str7 + ", " + str8 + ", " + str9);

            }
        }

        /// <summary>
        /// Logs the template offset related information
        /// </summary>
        /// <param name="templateRep">Template or TemplateReport object</param>
        private void ShowTemplateOffSetData(TemplateReport templateRep)
        {

            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - GET OFFSET DATA ");
            logService.WriteToFile("================================================================");

            double offsetDist = 0.0;
            long lowerSectionCount = 0, upperSectionCount = 0;
            try
            {
                templateRep.GetOffsetData(out offsetDist, out lowerSectionCount, out upperSectionCount);
            }
            catch (Exception ce)
            {
                string str = ce.ToString();
            }

            logService.WriteToFile(" ");
            logService.WriteToFile("Template Offset Distance      :  " + RoundedToPrecision(offsetDist, Precision));
            logService.WriteToFile("Template Lower Section Count  :  " + lowerSectionCount);
            logService.WriteToFile("Template Upper Section Count  :  " + upperSectionCount);
            logService.WriteToFile(" ");

        }

        /// <summary>
        /// Logs the templates marks data
        /// </summary>
        /// <param name="templateRep">Template or TemplateReport object</param>
        private void ShowTemplateMarksData(TemplateReport templateRep)
        {

            logService.WriteToFile("");

            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - GET MARKS ");
            logService.WriteToFile("================================================================");

            logService.WriteToFile(" XFrame Marks   : ");
            ReadOnlyCollection<ManufacturingGeometry> marksCol = null;
            string str = "";
            try
            {
                marksCol = templateRep.GetMarks(ManufacturingGeometryType.XFrameMark);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (marksCol != null)
            {
                WriteMarksInfoToLog(marksCol);
            }


            logService.WriteToFile(" YFrame Marks : ");
            try
            {
                marksCol = templateRep.GetMarks(ManufacturingGeometryType.YFrameMark);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (marksCol != null)
            {
                WriteMarksInfoToLog(marksCol);
            }


            logService.WriteToFile(" ZFrame Marks   : ");
            try
            {
                marksCol = templateRep.GetMarks(ManufacturingGeometryType.ZFrameMark);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (marksCol != null)
            {
                WriteMarksInfoToLog(marksCol);
            }

            logService.WriteToFile(" Template Marks    : ");
            try
            {
                marksCol = templateRep.GetMarks(ManufacturingGeometryType.TemplateMark);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            if (marksCol != null)
            {
                WriteMarksInfoToLog(marksCol);
            }

            logService.WriteToFile(" Seam Marks        : ");
            try
            {
                marksCol = templateRep.GetMarks(ManufacturingGeometryType.SeamMark);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }
            if (marksCol != null)
            {
                WriteMarksInfoToLog(marksCol);
            }
        }

        private void WriteMarksInfoToLog(ReadOnlyCollection<ManufacturingGeometry> marksCol)
        {

            if (marksCol == null)
            {
                logService.WriteToFile("  Number of Marks                           : " + 0);
                logService.WriteToFile(" ");
                return;
            }
            else
            {
                logService.WriteToFile("  Number of Marks                           : " + marksCol.Count);
                logService.WriteToFile(" ");

                foreach (ManufacturingGeometry mark in marksCol)
                {
                    BusinessObject markObj = mark.RepresentedEntity;
                    NamedItemHelper markName = new NamedItemHelper(markObj);
                    ComplexString3d curve = mark.Geometry;
                    Position start, end;
                    curve.EndPoints(out start, out end);
                    logService.WriteToFile("//  Mark Name                                 : " + markName.Name);
                    logService.WriteToFile("//  Mark Start Position                       : " + "(" + RoundedToPrecision(start.X, Precision) + ", " + RoundedToPrecision(start.Y, Precision) + ", " + RoundedToPrecision(start.Z, Precision) + ")");
                    logService.WriteToFile("//  Mark End Position                         : " + "(" + RoundedToPrecision(end.X, Precision) + ", " + RoundedToPrecision(end.Y, Precision) + ", " + RoundedToPrecision(end.Z, Precision) + ")");
                    logService.WriteToFile(" ");
                }
                logService.WriteToFile(" ");
            }

        }


        private void ShowTemplateDrawingInformation(TemplateReport templateRep)
        {

            logService.WriteToFile("================================================================");
            logService.WriteToFile(" TEMPLATE REPORT - DRAWING INFORMATION");
            logService.WriteToFile("================================================================");

            string str = "";

            //length of template                    
            double length = 0.0;
            try
            {
                length = templateRep.GetDistance(TemplateDistanceType.LowerEndOnTopLineAndUpperEndOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Lengh of Template                      : " + RoundedToPrecision(length, Precision));

            //height at length of template                    
            double heightAtLength = 0.0;
            try
            {
                heightAtLength = templateRep.GetHeight(length);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }
            logService.WriteToFile("Height at Lengh of Template            : " + RoundedToPrecision(heightAtLength, Precision));

            //height at lower seam
            double heightAtLowerSeam = 0.0;
            double lengthToLowerSeamOnTopLine = 0.0;
            try
            {
                lengthToLowerSeamOnTopLine = templateRep.GetDistance(TemplateDistanceType.LoweEndOnTopLineAndLowerSeamOnTopLine);
                heightAtLowerSeam = templateRep.GetHeight(lengthToLowerSeamOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Height at Lower Seam                   : " + RoundedToPrecision(heightAtLowerSeam, Precision));

            //Girth Along lower Seam To Mid Template
            double girthAlongLowerSeamToMidTemplate = 0.0;
            try
            {
                girthAlongLowerSeamToMidTemplate = templateRep.GetDistance(TemplateDistanceType.GirthBetweenLowerSeamAndReferenceLowerSeam);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Girth Along lower Seam To Mid Template : " + RoundedToPrecision(girthAlongLowerSeamToMidTemplate, Precision));

            //height at Upper seam
            double heightAtUpperSeam = 0.0;
            try
            {
                double lengthToUpperSeamOnTopLine = lengthToLowerSeamOnTopLine + templateRep.GetDistance(TemplateDistanceType.LowerSeamOnTopLineAndUpperSeamOnTopLine);
                heightAtUpperSeam = templateRep.GetHeight(lengthToUpperSeamOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Height at Upper Seam                   : " + RoundedToPrecision(heightAtUpperSeam, Precision));

            //Girth Along Upper Seam To Mid Template
            double girthAlongUpperSeamToMidTemplate = 0.0;
            try
            {
                girthAlongUpperSeamToMidTemplate = templateRep.GetDistance(TemplateDistanceType.GirthBetweenUpperSeamAndReferenceUpperSeam);

            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Girth Along Upper Seam To Mid Template : " + RoundedToPrecision(girthAlongUpperSeamToMidTemplate, Precision));

            //Height at Base Control Line 
            double heightAtBCL = 0.0;
            try
            {
                double lengthToBCLOnTopLine = lengthToLowerSeamOnTopLine + templateRep.GetDistance(TemplateDistanceType.LowerSeamOnTopLineAndBaseControlOnTopLine);
                heightAtBCL = templateRep.GetHeight(lengthToBCLOnTopLine);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Height at Base Control Line            : " + RoundedToPrecision(heightAtBCL, Precision));

            //Girth Along BCL To Mid Template 
            double girthAlongBCLToMidTemplate = 0.0;
            try
            {
                girthAlongBCLToMidTemplate = templateRep.GetDistance(TemplateDistanceType.GirthBetweenBaseControlAndReferenceBaseControl);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Girth Along BCL To Mid Template        : " + RoundedToPrecision(girthAlongBCLToMidTemplate, Precision));

            //Template Angle
            TemplateSide templateSide = TemplateSide.Aft;
            TemplateProjectionSide projectionSide = TemplateProjectionSide.PositiveX;
            double angle = 0.0;
            try
            {
                angle = templateRep.GetAngle(TemplateAngleType.Default, 0.2, out templateSide, out projectionSide);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile("Template Angle                         : " + RoundedToPrecision((angle) * 180 / Math.PI, Precision));

            //S_Value
            try
            {
                logService.WriteToFile("S_Value                                : " + RoundedToPrecision(templateRep.GetDistance(TemplateDistanceType.AftSeamOnTopLineAndBaseControlOnTopLine), Precision));
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            logService.WriteToFile(" ");
            logService.WriteToFile(" ");

        }


        private void ShowTemplateIntersectionInformation(Template mainTemplate)
        {

            string str = "";
            TemplateReport templateRep = new TemplateReport(mainTemplate);

            ReadOnlyCollection<Template> intersecTempColl = null;
            try
            {
                intersecTempColl = templateRep.GetIntersectingTemplates(TemplatePositionType.UpperSeam);
            }
            catch (Exception ce)
            {
                str = ce.ToString();
            }

            int numIntersectionTemplates = intersecTempColl.Count;            

            if ( numIntersectionTemplates > 0 )
            {

                logService.WriteToFile(" ");
                logService.WriteToFile("================================================================");
                logService.WriteToFile(" TEMPLATE REPORT - Templates Intersection Information ");
                logService.WriteToFile("================================================================");
                logService.WriteToFile("");
                
                logService.WriteToFile("Number of Intersection Templates: " + numIntersectionTemplates);

                int nCounter = 0;

                foreach (Template intersectedTemplate in intersecTempColl)
                {
                    TemplateReport intersectedTemplateRep = new TemplateReport(intersectedTemplate);
                    NamedItemHelper intersectedTemplateName = new NamedItemHelper(intersectedTemplate);

                    logService.WriteToFile("Template Name: " + intersectedTemplateName.Name);
                    TemplateIntersectionInfo tempIntersecInfo = null;

                    try
                    {
                        tempIntersecInfo = templateRep.GetIntersectionInfo(intersectedTemplate);

                    }
                    catch (Exception ce)
                    {
                        str = ce.ToString();
                    }

                    Position bottomPos = null;
                    Position topPos = null;
                    if (tempIntersecInfo != null )
                    {
                        bottomPos = tempIntersecInfo.BottomIntersectionPosition;

                        if (null != bottomPos)
                        {
                            logService.WriteToFile("BottomLineIntersectionPoint  : " + RoundedToPrecision(bottomPos.X, Precision) + ", " + RoundedToPrecision(bottomPos.Y, Precision) + ", " + RoundedToPrecision(bottomPos.Z, Precision));
                        }
                        
                        topPos = tempIntersecInfo.TopIntersectionPosition;

                        if (null != topPos)
                        {
                            logService.WriteToFile("TopLineIntersectionPoint  : " + RoundedToPrecision(topPos.X, Precision) + ", " + RoundedToPrecision(topPos.Y, Precision) + ", " + RoundedToPrecision(topPos.Z, Precision));
                        }
                        logService.WriteToFile("");

                    }                    

                    nCounter++;
                }
            }
        }
                
        internal static double RoundedToPrecision(double number, long numDigits)
        {
            /* This function rounds a number to a specified number of decimal
             places. 0.5 is rounded up.
             The default behavior of VB is as follows, which this function overrides:
             When you convert a decimal value to an integer value, VBA rounds the number
             to an integer value. How it rounds depends on the value of the digit
             immediately to the right of the decimal place.  Digits less than 5 are
             rounded down, while digits greater than 5 are rounded up.
             If the digit is 5, then it's rounded down if the digit immediately to the
             left of the decimal place is even, and up if it's odd.
             When the digit to be rounded is a 5, the result is always an even integer. */

            double power = Math.Pow(10, numDigits);
            int intSgn = Math.Sign(number);

            number = Math.Abs(number);
            //Do the major calculation.
            double temp = number * power + 0.50000000001;

            int tempint = (int)temp;

            //Finish the calculation.
            double value = intSgn * (tempint / power);
            return value;

        }
        
        private static string BoundaryName(BusinessObject boundary)
        {

            if (boundary == null)
            {
                return "Free Edge";
            }
            else if (boundary.SupportsInterface("IJNamedItem"))
            {
                NamedItemHelper boundaryName = new NamedItemHelper(boundary);
                return boundaryName.Name;
            }
            else if (boundary.SupportsInterface("IJStructPort"))
            {
                return "Profile Face";
            }
            else
            {
                return "";
            }
        }


        #endregion Private Methods


    }
}
