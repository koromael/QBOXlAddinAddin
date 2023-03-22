

using Intuit.Ipp.Data;
using System;


#region Get XML From Report object 

/*==================================================
=            Get XML From Report object            =
==================================================*/
class XmlConverter
{
    public string getReportInXML(Report report)
    {
        string strreport = "";
        strreport += "<Report xmlns=\"http://schema.intuit.com/finance/v3\">";
        //strreport += getReportHeader(report.Header);
        //strreport += getReportColumns(report.Columns);
        strreport += getRows(report.Rows);
        strreport += "</Report>";
        return strreport;
    }

    private string getReportHeader(ReportHeader header)
    {
        string strheader = "";
        strheader += "<Header>";
        strheader += "<Time>" + String.Format("{0:yyyy-MM-ddTH:mm:sszzz}", header.Time) + "</Time>";
        strheader += "<ReportName>" + header.ReportName + "</ReportName>";
        strheader += "<DateMacro>" + header.DateMacro + "</DateMacro>";
        strheader += "<ReportBasis>" + header.ReportBasis + "</ReportBasis>";
        strheader += "<StartPeriod>" + header.StartPeriod + "</StartPeriod>";
        strheader += "<EndPeriod>" + header.EndPeriod + "</EndPeriod>";
        strheader += "<SummarizeColumnsBy>" + header.SummarizeColumnsBy + "</SummarizeColumnsBy>";
        strheader += "<Currency>" + header.Currency + "</Currency>";
        foreach (var option in header.Option)
        {
            strheader += "<Option>";
            strheader += "<Name>" + option.Name + "</Name>";
            strheader += "<Value>" + option.Value + "</Value>";
            strheader += "</Option>";
        }
        strheader += "</Header>";
        return strheader;
    }


    private string getReportColumns(Column[] columns)
    {
        string strcolumn = "";
        strcolumn += "<Columns>";
        foreach (Column column in columns)
        {
            strcolumn += "<Column>";
            strcolumn += " <ColTitle>" + Convert.ToString(column.ColTitle) + "</ColTitle>";
            strcolumn += " <ColType>" + column.ColType + "</ColType>";
            foreach (var meta in column.MetaData)
            {
                strcolumn += "<MetaData>";
                strcolumn += "<Name>" + meta.Name + "</Name>";
                strcolumn += "<Value>" + meta.Value + "</Value>";
                strcolumn += "</MetaData>";
            }
            strcolumn += "</Column>";
        }
        strcolumn += "</Columns>";
        return strcolumn;
    }
    private string getRows(Row[] rows)
    {
        string strrows = "";
        strrows += " <Rows>";
        foreach (Row row in rows)
        {
            if (row.type.ToString() == "Section")                    ///It contains more Rows
            {
                // if type of row is data then call it
                strrows += getRowSection(row);
            }
            else if (row.type.ToString() == "Data")
            {
                //Otherwise call getRowSection
                strrows += getRowData(row);
            }
        }
        strrows += " </Rows>";
        return strrows;
    }
    private string getRowSection(Row row)
    {
        string strrowsection = "";
        strrowsection += " <Row type=\"Section\"";
        if (row.group != null)
        {
            strrowsection += " group=\"" + row.group + "\"";
        }
        strrowsection += ">";
        string type = row.AnyIntuitObjects[0].GetType().ToString();
        object[] objs = (object[])row.AnyIntuitObjects;
        foreach (var obj in objs)
        {
            string objtype = obj.GetType().ToString();
            switch (objtype)
            {
                case "Intuit.Ipp.Data.Header":
                    Header header = (Header)obj;
                    strrowsection += getHeader(header);
                    break;
                case "Intuit.Ipp.Data.Rows":
                    Rows rows = (Rows)obj;
                    strrowsection += getRows(rows.Row);
                    break;
                case "Intuit.Ipp.Data.Summary":
                    Summary summary = (Summary)obj;
                    strrowsection += getSummary(summary);
                    break;
            }
        }
        strrowsection += " </Row>";
        return strrowsection;
    }
    private string getSummary(Summary summary)
    {
        string strsummary = "";
        strsummary += "<Summary>";
        ColData[] coldata = (ColData[])summary.ColData;
        strsummary += getColData(coldata);
        strsummary += "</Summary>";
        return strsummary;
    }
    private string getRowData(Row row)
    {
        string strrowdata = "";
        strrowdata += " <Row type=\"Data\">";
        ColData[] coldata = (ColData[])row.AnyIntuitObjects[0];
        strrowdata += getColData(coldata);
        strrowdata += "</Row>";
        return strrowdata;
    }
    private string getHeader(Header header)
    {
        string strheader = "";
        strheader += "   <Header>";
        ColData[] coldata = (ColData[])header.ColData;
        strheader += getColData(coldata);
        strheader += "   </Header>";
        return strheader;
    }
    private string getColData(ColData[] coldata)
    {
        string strcoldata = "";
        foreach (ColData col in coldata)
        {
            //Generate the coldata string here.
            string value = col.value;
            value = value.Replace("&", "&amp;");
            strcoldata += "<ColData value=\"" + value + "\"";
            if (col.id != null)
            {
                strcoldata += " id=\"" + col.id + "\"";
            }
            strcoldata += " /> ";
        }
        return strcoldata;
    }
    /*=====  End of Get XML From Report object  ======*/
    #endregion  End of Get XML From Report object

}