/*---------------------------------------------------------------------+\
|	Copyright 2019 Hexagon PPM									        |
|	All Rights Reserved													|
|																		|
|	Including software, file formats, and audio-visual displays;		|
|	may only be used pursuant to applicable software license			|
|	agreement; contains confidential and proprietary information of		|
|	Hexagon PPM and/or third parties which is protected by copyright	|
|	and trade secret law and may not be provided or otherwise made		|
|	available without proper authorization.								|
|																		|
|	Unpublished -- rights reserved under the Copyright Laws of the		|
|	United States.														|
|																		|
|	Hexagon PPM         												|
|	Huntsville, Alabama	  35894-0001									|
\+---------------------------------------------------------------------*/

/*---------------------------------------------------------------------+\
|   CQueryInterpreter.cs : Implementation of CQueryInterpreter
|
|	Purpose:
|	File Custodian:		Garam Han
\+---------------------------------------------------------------------*/

/*---------------------------------------------------------------------+\
|	Revision History:					(most recent entries first)
|
|	03-Sep-2019 Garam Han
|		DM-CP-353905  Needs to customize Weld Labels  
|
|	24-Apr-2019 Garam Han
|		DI-CP-347726 Implement fix for Coverity issues - DwgWeldLabel CQueryInterpreter.cs 
|
|	25-Feb-2019 Garam Han
|		Initial Revision
\+---------------------------------------------------------------------*/

using System;
using System.Data;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Reports.Middle;
using IMSRelation;

namespace DwgWeldLabelQI
{    
    public class CQueryInterpreter : QueryInterpreter
    {
        const string MODULE = "DwgWeldLabelQI.CQueryInterpreter";

        private DataTable m_DataTable;

        #region IJQueryInterpreter Members

        public override DataTable Execute(string Action, string Argument)
        {
            try
            {
                // Create a record DataColumn list that will contain all the column DataColumns
                List<Column> FieldColumnList = new List<Column>();

                Column recFieldColumnOID = new Column("OID", typeof(System.String));
                FieldColumnList.Add(recFieldColumnOID);

                Column recFieldColumnWeldLegLength = new Column("Weld Leg Length", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnWeldLegLength);

                Column recFieldColumnWeldThroatThickness = new Column("Weld Throat Thickness", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnWeldThroatThickness);

                Column recFieldColumnWeldThroatThicknessSqrt = new Column("Weld Throat Thickness_Sqrt", typeof(System.Double));
                FieldColumnList.Add(recFieldColumnWeldThroatThicknessSqrt);

                Column recFieldColumnFilletMeasureMethod = new Column("Fillet Measure Method", typeof(System.String));
                FieldColumnList.Add(recFieldColumnFilletMeasureMethod);

                // create and initialize the Data Table with the Columns
                m_DataTable = InitializeDataTable(FieldColumnList);

                // return empty tabel when EvaludateOnly is True
                if (EvaluateOnly)
                {
                    return m_DataTable;
                }

                foreach (BusinessObject oBOTemp in InputObjects)
                {
                    PropertyValueCodelist codelistMeasureMethod = null;
                    BusinessObject oBO = oBOTemp;
                    string measureMethodVal = "";

                    if (oBO.SupportsInterface("IJUASmartWeld")) // Physical Connection
                    {
                        codelistMeasureMethod = (PropertyValueCodelist)oBO.GetPropertyValue("IJUASmartWeld", "FilletMeasureMethod");
                    }
                    else if (oBO.SupportsInterface("IJPlnJoint")) // Planning Joint
                    {
                        codelistMeasureMethod = (PropertyValueCodelist)oBO.GetPropertyValue("IJPlnJointProps", "WeldFilletMeasureMethod");

                        IJDAssocRelation objAssocRelation = (IJDAssocRelation)(COMConverters.ConvertBOToCOMBO(oBO));
                        if (objAssocRelation != null)
                        {
                            IMSRelation.IJDTargetObjectCol objTargetCol = objAssocRelation.get_CollectionRelations("IJPlnJoint", "PlnJointPhysConn_DEST") as IMSRelation.IJDTargetObjectCol;
                            if (objTargetCol != null)
                            {
                                if (objTargetCol.Count > 0)
                                    oBO = COMConverters.ConvertCOMBOToBO(objTargetCol.get_Item(1) as object);
                            }
                        }
                    }

                    if (codelistMeasureMethod != null)
                        measureMethodVal = codelistMeasureMethod.ToString();

                    double legLengthVal = 0.0, ThroatThicknessVal = 0.0;
                    if (oBO.SupportsInterface("IJWeldingSymbol")) // get values from IJWeldingSymbol
                    {
                        PropertyValueDouble dblActualLegLength = null;
                        PropertyValueDouble dblNominalThroatThickness = null;
                        if (String.Compare(Action.ToUpper(), "PRIMARY SIDE") == 0)
                        {
                            dblActualLegLength = (PropertyValueDouble)oBO.GetPropertyValue("IJWeldingSymbol", "PrimarySideActualLegLength");
                            dblNominalThroatThickness = (PropertyValueDouble)oBO.GetPropertyValue("IJWeldingSymbol", "PrimarySideNominalThroatThickness");
                        }
                        else if (String.Compare(Action.ToUpper(), "SECONDARY SIDE") == 0)
                        {
                            dblActualLegLength = (PropertyValueDouble)oBO.GetPropertyValue("IJWeldingSymbol", "SecondarySideActualLegLength");
                            dblNominalThroatThickness = (PropertyValueDouble)oBO.GetPropertyValue("IJWeldingSymbol", "SecondarySideNominalThroatThickness");
                        }

                        if (null != dblActualLegLength)
                        {
                            if (null != dblActualLegLength.PropValue)
                            {
                                legLengthVal = (double)dblActualLegLength.PropValue;
                            }
                        }

                        if (null != dblNominalThroatThickness)
                        {
                            if (null != dblNominalThroatThickness.PropValue)
                            {
                                ThroatThicknessVal = (double)dblNominalThroatThickness.PropValue;
                            }
                        }
                    }

                    // Set Data Row
                    DataRow currentRow = m_DataTable.NewRow();

                    currentRow.SetField("OID", oBO.ObjectID.ToString());
                    if (legLengthVal == 0.0)
                        currentRow.SetField("Weld Leg Length", 0); 
                    else
                        currentRow.SetField("Weld Leg Length", legLengthVal);
                    currentRow.SetField("Weld Throat Thickness", ThroatThicknessVal);
                    currentRow.SetField("Weld Throat Thickness_Sqrt", ThroatThicknessVal * Math.Sqrt(2));
                    currentRow.SetField("Fillet Measure Method", measureMethodVal);

                    // Set Delegated DataTable
                    System.Data.DataTable DelegatedDataTable = ExecuteDelegatedQuery(oBO);

                    if (DelegatedDataTable != null)
                    {
                        if (DelegatedDataTable.Rows.Count >= 1)
                        {
                            DataRow row = DelegatedDataTable.Rows[0];

                            // merge the column value pair to data table, excluding the oid column
                            foreach (DataColumn col in row.Table.Columns)
                            {
                                if (col.ColumnName.ToLower() != "oid")
                                {
                                    currentRow.SetField(col.ColumnName, row[col.ColumnName]);
                                }
                            }
                        }
                    }

                    m_DataTable.Rows.Add(currentRow);
                }

                return m_DataTable;
            }
            catch (Exception e)
            {
                throw new Exception("Failed in Execute because " + e.Message);
            }
        }

        #endregion
    }
}
