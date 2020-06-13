using NUnit.Framework;
using ServiceStatus.Helpers;
using System.Data;
using System.Text;

namespace ServiceStatus
{
    public class ProcessTest
    {
        [Test]
        public void TestServiceRunning()
        {
            ProcessExecutorHelper p = new ProcessExecutorHelper();
            EmailHelper e = new EmailHelper();
            HtmlHelper htmlHelper = new HtmlHelper();

            // Get Excel Datatable
            DataTable excelDataTable = ExcelHelper.ExcelRead("Inputs",ExcelHelper.filePath);

            // Return Result Datatable
            p.PerformProcessAndResultDataGenerate(excelDataTable);

            // Generate Result
            StringBuilder htmlResult =htmlHelper.GenerateHtmlResult(ProcessExecutorHelper.MachineServicesList);

            // Email Result
            e.SendEmail(htmlResult);
        }
    }
}
