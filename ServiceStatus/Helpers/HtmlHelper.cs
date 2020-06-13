using ServiceStatus.Property;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using static System.Configuration.ConfigurationManager;

namespace ServiceStatus.Helpers
{
    public class HtmlHelper
    {
        public StringBuilder GenerateHtmlResult(List<MachineData> machineServicesList)
        {
            bool isHeader = true;
            var services = string.Empty;
            var machines = string.Empty;

            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<p>Ran from Machine : " + Environment.MachineName + "</p>");
            strHTMLBuilder.Append("<p>Machine User Name : " + Environment.UserName + " </p>");
            strHTMLBuilder.Append("<p>User Id used To Access Other System : " + AppSettings["userName"] + " </p>");
            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:smaller'>");

            string headerStyle = "bgcolor = \"LightGray\"";
            string isAsExpectedStyle = "bgcolor = \"Green\"";// Green
            string isNotAsExpectedStyle = "bgcolor = \"Red\"";// Red

            foreach (var machine in machineServicesList)
            {
                int unskippedServiceCount = 0;
                if (machine.IsMachineSkipped)
                    continue;
                if (isHeader)
                {
                    strHTMLBuilder.Append("<tr " + headerStyle + ">");
                    strHTMLBuilder.Append("<td>");
                    strHTMLBuilder.Append("Machine Name");
                    strHTMLBuilder.Append("</td>");
                    foreach (var service in machine.services)
                    {
                        if (!service.IsServiceSkipped)
                        {
                            strHTMLBuilder.Append("<td>");
                            strHTMLBuilder.Append(service.Service);
                            strHTMLBuilder.Append("</td>");
                        }
                    }
                    strHTMLBuilder.Append("</tr>");
                    isHeader = false;
                }

                strHTMLBuilder.Append("<tr>");
                strHTMLBuilder.Append("<td>");
                strHTMLBuilder.Append(machine.MachineName);
                strHTMLBuilder.Append("</td>");

                if (machine.ErrorMessage == null)
                {
                    foreach (var service in machine.services)
                    {
                        if (!service.IsServiceSkipped)
                        {
                            strHTMLBuilder.Append("<td " + (service.IsServiceCheckSuccess ? isAsExpectedStyle : isNotAsExpectedStyle) + ">");
                            strHTMLBuilder.Append(service.CurrentServiceStatus);
                            strHTMLBuilder.Append("</td>");
                        }
                    }
                }
                else
                {
                    foreach (var service in machine.services)
                        if (!service.IsServiceSkipped)
                            unskippedServiceCount++;
                    strHTMLBuilder.Append("<td colspan=" + unskippedServiceCount + " " + isNotAsExpectedStyle + ">");
                    strHTMLBuilder.Append(machine.ErrorMessage);
                    strHTMLBuilder.Append("<td>");
                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</table>");

            foreach (var machine in machineServicesList)
            {
                if (machine.IsMachineSkipped)
                    machines += machine.MachineName + ",";
                else
                    foreach (var service in machine.services)
                    {
                        if (service.IsServiceSkipped)
                            services += service.Service + ",";

                        //if (service.IsServiceCheckSuccess)
                        if (service.IsStartStopSuccess)
                            strHTMLBuilder.Append("<p>" + service.StartStopMessage + " for Machine " + machine.MachineName + "</p>");
                    }
                if (services != string.Empty)
                    strHTMLBuilder.Append("<p>Service " + services + " skipped for " + machine.MachineName + "</p>");
            }
            if (machines != string.Empty)
                strHTMLBuilder.Append("<p>Service check skipped for Machine " + machines + "</p>");

            return strHTMLBuilder;
        }

        public static string ExportDatatableToHtml(DataTable dt)
        {
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");
            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:smaller'>");

            strHTMLBuilder.Append("<tr >");
            foreach (DataColumn myColumn in dt.Columns)
            {
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(myColumn.ColumnName);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append("</td>");
                }
                strHTMLBuilder.Append("</tr>");
            }
            //Close tags.  
            strHTMLBuilder.Append("</table>");
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");
            string Htmltext = strHTMLBuilder.ToString();
            return Htmltext;
        }
    }
}
