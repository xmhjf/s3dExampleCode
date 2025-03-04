using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Reports.Middle;
using System.Data;
using Ingr.SP3D.Common.Middle.Services;
using System.Diagnostics;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Support.Content.BOMLabels
{
    
    /// <summary>
    /// This is the custom label QueryInterpreter. The user
    /// should override the Run method.
    /// </summary>
    public class Utility_TWO_HOLE_PLATE_BOMLabel : QueryInterpreter
    {
        private DataTable m_DataTable;

        /// <summary>
        /// constructor
        /// </summary>
        public Utility_TWO_HOLE_PLATE_BOMLabel()
        {

        }

        /// <summary>
        /// This is the main method that will get called from the Runtime framework.
        /// </summary>
        /// <param name="Command"></param>
        /// <param name="Argument"></param>
        public override DataTable Execute(string action, string argument)
        {
            try
            {

                // Create a record DataColumn list that will contain all the
                // column DataColumns
                List<Column> FieldColumnList = new List<Column>();

                Column recFieldColumnWidth = new Column("Width", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnWidth);

                Column recFieldColumnDepth = new Column("Depth", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnDepth);

                Column recFieldColumnHoleSize = new Column("HoleSize", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnHoleSize);

                Column recFieldColumnC = new Column("C", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnC);

                Column recFieldColumnThickness = new Column("Thickness", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnThickness);


                // create and initialize the Data Table with the Columns ( user added 
                // columns and returned properties from the Delegated Query)
                m_DataTable = InitializeDataTable(FieldColumnList);

                // if the Evaluate property is set to true that means we need to
                // return the data table with only columns names to the designer
                // so that the control can display the column names at the design time
                // The Query Builder calls the Execute with Evaluate set to true
                if (EvaluateOnly)
                {
                    return m_DataTable;
                }

                // Get the first object in the input objects
                BusinessObject oBO = InputObjects.ElementAt<BusinessObject>(0);

                double Width = 0.0, Depth = 0.0, HoleSize = 0.0, C = 0.0, Thickness = 0.0;

                if (oBO is Support.Middle.SupportComponent)
                {
                    Support.Middle.SupportComponent supportComp = (Support.Middle.SupportComponent)oBO;

                    Width = (double)((PropertyValueDouble)supportComp.GetPropertyValue("IJOAHgrUtility_TWO_HOLE_PLATE", "WIDTH")).PropValue;
                    Depth = (double)((PropertyValueDouble)supportComp.GetPropertyValue("IJOAHgrUtility_TWO_HOLE_PLATE", "DEPTH")).PropValue;
                    HoleSize = (double)((PropertyValueDouble)supportComp.GetPropertyValue("IJOAHgrUtility_TWO_HOLE_PLATE", "DEPTH")).PropValue;
                    C = (double)((PropertyValueDouble)supportComp.GetPropertyValue("IJOAHgrUtility_TWO_HOLE_PLATE", "C")).PropValue;

                    Part part = (Part)supportComp.GetRelationship("madeFrom", "part").TargetObjects[0];
                    Thickness = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_TWO_HOLE_PLATE", "THICKNESS")).PropValue;
                                        
                }

                // data row to hold the data for current object
                DataRow currentRow = m_DataTable.NewRow();

                currentRow.SetField("Width", Width);
                currentRow.SetField("Depth", Depth);
                currentRow.SetField("HoleSize", HoleSize);
                currentRow.SetField("C", C);
                currentRow.SetField("Thickness", Thickness);

                // Execute the Delegated Query.
                System.Data.DataTable DelegatedDataTable = ExecuteDelegatedQuery(oBO);

                // add the new row with the value  to the record set. Below code is especially for COM labels. 
                if (DelegatedDataTable != null)
                {
                    if (DelegatedDataTable.Rows.Count >= 1)
                    {
                        DataRow row = DelegatedDataTable.Rows[0];

                        // merge the column value pair to data table
                        foreach (DataColumn col in row.Table.Columns)
                        {
                            currentRow.SetField(col.ColumnName, row[col.ColumnName]);
                        }
                    }
                }
                // add the new row to the master data table
                m_DataTable.Rows.Add(currentRow);

                return m_DataTable;
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(0, typeof(Utility_TWO_HOLE_PLATE_BOMLabel).FullName, e.ToString(), string.Empty, (new StackTrace()).GetFrame(1).GetMethod().Name, string.Empty, string.Empty, -1);
                throw;
            }
        }

    }


}
