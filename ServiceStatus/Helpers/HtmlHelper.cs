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
            var servicesSkipped = string.Empty;
            var machinesSkipped = string.Empty;
            string servicesStarted = string.Empty;
            string servicesStopped = string.Empty;

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
                    strHTMLBuilder.Append("<th>");
                    strHTMLBuilder.Append("Machine Name");
                    strHTMLBuilder.Append("</th>");
                    foreach (var service in machine.services)
                    {
                        if (!service.IsServiceSkipped)
                        {
                            strHTMLBuilder.Append("<th>");
                            strHTMLBuilder.Append(service.Service);
                            strHTMLBuilder.Append("</th>");
                        }
                    }
                    strHTMLBuilder.Append("</tr>");
                    isHeader = false;
                }

                strHTMLBuilder.Append("<tr>");
                strHTMLBuilder.Append("<td><b>");
                strHTMLBuilder.Append(machine.MachineName);
                strHTMLBuilder.Append("</b></td>");

                if (machine.ErrorMessage == null)
                {
                    foreach (var service in machine.services)
                        if (!service.IsServiceSkipped)
                        {
                            strHTMLBuilder.Append("<td " + (service.IsServiceStatusAsExpected ? isAsExpectedStyle : isNotAsExpectedStyle) + ">");
                            strHTMLBuilder.Append(service.CurrentServiceStatus);
                            strHTMLBuilder.Append("</td>");
                        }
                }
                else
                {
                    foreach (var service in machine.services)
                        if (!service.IsServiceSkipped)
                            unskippedServiceCount++;
                    strHTMLBuilder.Append("<td colspan=" + unskippedServiceCount + " " + isNotAsExpectedStyle + ">");
                    strHTMLBuilder.Append(machine.ErrorMessage);
                    strHTMLBuilder.Append("</td>");
                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</table>");

            foreach (var machine in machineServicesList)
            {
                string machineName = machine.MachineName;
                if (machine.IsMachineSkipped)
                {
                    machinesSkipped += machine.MachineName + ",";
                    continue;
                }
                else if (machine.ErrorMessage != null)
                    continue;
                else
                    foreach (var service in machine.services)
                    {
                        if (service.IsServiceSkipped)
                            servicesSkipped += service.Service + ",";

                        //if (service.IsServiceCheckSuccess)
                        if (service.IsStartStopSuccess)
                        {
                            if (service.IsStarted)
                                servicesStarted += service.Service + ",";

                            if (service.IsStopped)
                                servicesStopped += service.Service + ",";
                        }
                        else { strHTMLBuilder.Append(service.StartStopMessage); }
                    }
                if (servicesSkipped != string.Empty)
                    strHTMLBuilder.Append("<p>Service <b>" + servicesSkipped.TrimEnd(',') + " skipped </b>for<b> " + machineName + "</b></p>");
                if (servicesStarted != string.Empty)
                    strHTMLBuilder.Append("<p><b>" + servicesStarted.TrimEnd(',') + "</b> service <b>Started Successfully </b>" + " for Machine <b>" + machine.MachineName + "</b></p>");
                if (servicesStopped != string.Empty)
                    strHTMLBuilder.Append("<p><b>" + servicesStopped.TrimEnd(',') + "</b> service <b> Stopped Successfully </b>" + " for Machine <b>" + machine.MachineName + "</b></p>");
            }
            if (machinesSkipped != string.Empty)
                strHTMLBuilder.Append("<p>Service check <b>skipped</b> for Machine <b>" + machinesSkipped.TrimEnd(',') + "</b></p>");
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
