using NUnit.Framework;
using System.Collections.Generic;

namespace ServiceStatus
{
    [TestFixture]
    class ServiceTest
    {
        //[Test]
        public void testServiceRunning()
        {
            ProcessExecutorHelper p = new ProcessExecutorHelper();
            FileHelper f = new FileHelper();
            EmailHelper e = new EmailHelper();
            //string pathOfBatchFile = @"C:\Users\SumitThapar\Documents\ServiceFile\servicestatus.bat";
            //string pathOfOutputFile = @"C:\Users\SumitThapar\Documents\ServiceFile\Output.txt";

            // Get Services List
            //var servicesList = ExcelHelper.xlReadSpecificColoumn(3,2,5,1);
            //Dictionary<string, string> Company_Details = exHelper.GetDataIntoDictionary(2, 1, 2);
            //var serviceStatus=p.GetServiceStatus(servicesList);
           // p.LoginToMachine();
           // p.GetServiceStatusContent(ExcelHelper.filePath,"Sheet3");

            //p.ExecuteBatchFile(pathOfBatchFile);
            //var content = f.ReadFile(pathOfOutputFile);
            //e.SendEmail(serviceStatus);
        }
    }
}
