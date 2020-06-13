using System;
using System.Collections.Generic;
using System.Data;
using static System.Configuration.ConfigurationManager;
using System.Management;
using System.ServiceProcess;
using ServiceStatus.Property;

namespace ServiceStatus
{
    public class ProcessExecutorHelper
    {
        public static List<MachineData> MachineServicesList = null;
        public static bool LoginSuccess = false;

        public ManagementScope LoginToMachine(string machineAddress, out bool isLoginSuccess)
        {
            MachineData machineData = new MachineData();
            ManagementScope scope = null;
            try
            {
                string machineIp = GetIpAddressFromMachineAddress(machineAddress);
                string IP = $@"\\{machineIp}\root\cimv2"; // remote IP  
                string username = AppSettings["userName"]; // remote username  
                string password = AppSettings["password"]; // remote password  
                ConnectionOptions connectoptions = new ConnectionOptions();
                connectoptions.Username = username;
                connectoptions.Password = password;
                machineData.MachineName = machineAddress;
                machineData.MachineIp = machineIp;
                scope = new ManagementScope(IP, connectoptions);
                scope.Connect();
                LoginSuccess = true;
                try
                {
                    // Check if Services are accessible
                    ServiceController[] services = ServiceController.GetServices(scope.Path.Server);
                }
                catch (Exception ex)
                {
                    machineData.ErrorMessage = ex.Message;
                    LoginSuccess = false;
                }
            }
            catch (Exception e)
            {
                LoginSuccess = false;
                machineData.ErrorMessage = e.Message;
            }
            machineData.isSuccess = LoginSuccess;
            MachineServicesList.Add(machineData);
            isLoginSuccess = LoginSuccess;
            return scope;
        }

        public string GetIpAddressFromMachineAddress(string machineName)
        {
            var ServersDic = ExcelHelper.GetDataIntoDictionary("Servers", 1, 2);
            foreach (KeyValuePair<string, string> server in ServersDic)
                if (server.Key == machineName)
                    return server.Value;
            return null;
        }

        public void StartStopServices(IDictionary<string, string> serviceStatusExpected, ManagementScope scope = null)
        {
            ServiceController sc = null;
            foreach (KeyValuePair<string, string> val in serviceStatusExpected)
            {
                if (scope != null)
                    sc = new ServiceController(val.Key, scope.Path.Server);
                else
                    sc = new ServiceController(val.Key);

                var currentStatus = sc.Status.ToString();
                string expectedStatus = val.Value;

                if (currentStatus == expectedStatus)
                    continue;
                else
                {
                    if (expectedStatus == ServiceControllerStatus.Running.ToString())
                    {
                        // sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running);
                    }
                    else
                    {
                        // sc.Stop();
                        sc.WaitForStatus(ServiceControllerStatus.Stopped);
                    }
                }
            }
        }

        public void PerformProcessAndResultDataGenerate(DataTable excelDatatable)
        {
            DataTable resultDataTable = new DataTable();
            int machineCounter = 0;
            MachineServicesList = new List<MachineData>();

            // Iterate Webservers
            for (int rowCount = 2; rowCount < excelDatatable.Rows.Count; rowCount++)
            {
                ManagementScope scope = null;
                DataRow row = excelDatatable.Rows[rowCount];
                MachineData md = new MachineData();
                List<ServicesData> servicesFromExcel = null;
                bool isLoginSuccess = false;

                // Login to Webserver
                bool skipMachine = Convert.ToBoolean(row[1]);
                if (!skipMachine)
                {
                    md.IsMachineSkipped = false;
                    md.MachineName = row[0].ToString();
                    scope = LoginToMachine(md.MachineName, out isLoginSuccess);
                    md.isSuccess = isLoginSuccess;
                }
                else { md.IsMachineSkipped = true; continue; }

                //if (isLoginSuccess)
                //{
                // Iterate Services
                servicesFromExcel = GetServiceData(excelDatatable, md.MachineName);

                // Get Current Service Status
                ServiceController sc = GetCurrentServiceStatus(servicesFromExcel, scope);

                // Service ON/Off
                StartStopServices(servicesFromExcel, sc);
                //}

                // Add Service Data To Machine data                
                md.services = servicesFromExcel;
                MachineServicesList[machineCounter].IsMachineSkipped = md.IsMachineSkipped;
                MachineServicesList[machineCounter++].services = md.services;
                scope = null;
            }
        }

        public List<ServicesData> GetServiceData(DataTable excelDatatable, string machineName)
        {
            List<ServicesData> listServices = new List<ServicesData>();
            ServicesData sd = new ServicesData();
            for (int col = 2; col < excelDatatable.Columns.Count; col++)
            {
                DataColumn column = excelDatatable.Columns[col];
                if (Convert.ToBoolean(Convert.ToInt32(excelDatatable.Rows[1][col].ToString()))) continue;
                sd = new ServicesData();
                for (int row = 0; row < 2; row++)
                    if (excelDatatable.Rows[row][col].ToString() != "")
                    {
                        sd.Service = column.ColumnName;
                        if (row == 0) sd.ServiceName = excelDatatable.Rows[row][col].ToString();
                        if (row == 1) sd.ServiceFlag = Convert.ToBoolean(Convert.ToInt32(excelDatatable.Rows[row][col].ToString()));
                    }

                int machineRowCounter = 2;
                for (int i = machineRowCounter; i < excelDatatable.Rows.Count; i++)
                    // finding index of row of machine name
                    if (excelDatatable.Rows[i][0].ToString() == machineName)
                        machineRowCounter = i;
                sd.ExpectedServiceStatus = excelDatatable.Rows[machineRowCounter][col].ToString();
                sd.MachineIp = GetIpAddressFromMachineAddress(machineName);
                listServices.Add(sd);
            }
            return listServices;
        }

        public ServiceController GetCurrentServiceStatus(List<ServicesData> servicesFromExcel, ManagementScope scope = null)
        {
            ServiceController sc = null;
            bool isServiceCheckSuccess = false;
            foreach (var item in servicesFromExcel)
            {
                try
                {
                    if (!item.ServiceFlag)
                    {
                        if (scope != null)
                            sc = new ServiceController(item.ServiceName, scope.Path.Server);
                        else
                            sc = new ServiceController(item.ServiceName);
                        var currentStatus = sc.Status.ToString();
                        item.CurrentServiceStatus = currentStatus;
                        if (currentStatus.Contains("was not found on computer"))
                            isServiceCheckSuccess = false;
                        isServiceCheckSuccess = true;
                    }
                    else
                        // Skip service check
                        item.IsServiceSkipped = true;
                }
                catch (Exception ex)
                {
                    isServiceCheckSuccess = false;
                    item.CurrentServiceStatus = ex.Message;
                }
                item.IsServiceCheckSuccess = isServiceCheckSuccess;
            }
            return sc;
        }

        public void StartStopServices(List<ServicesData> servicesFromExcel, ServiceController sc)
        {
            bool isSuccess = false;
            string message = string.Empty;
            string currentServiceStatus = string.Empty;
            string expectedServiceStatus = string.Empty;
            string serviceName = string.Empty;
            string serviceDisplayName = string.Empty;
            foreach (var service in servicesFromExcel)
            {
                try
                {
                    currentServiceStatus = service.CurrentServiceStatus;
                    expectedServiceStatus = service.ExpectedServiceStatus;
                    serviceName = service.ServiceName;
                    serviceDisplayName = service.Service;

                    if (currentServiceStatus != null)
                        if (currentServiceStatus == expectedServiceStatus)
                            continue;
                        else
                        {
                            if (expectedServiceStatus == ServiceControllerStatus.Running.ToString())
                            {
                                sc.Start();
                                sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                                message = serviceDisplayName + "Service Started Successfully";
                            }
                            else if (expectedServiceStatus == ServiceControllerStatus.Stopped.ToString())
                            {
                                sc.Stop();
                                sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10));
                                message = serviceDisplayName + "Service Stopped Successfully";
                            }
                            isSuccess = true;
                        }
                }
                catch (Exception ex)
                {
                    isSuccess = false;
                    message = "Failure to change the service status from " + currentServiceStatus + " to " + expectedServiceStatus + " for Service " + serviceDisplayName + " Error -> " + ex.Message;
                }
                service.IsStartStopSuccess = isSuccess;
                service.StartStopMessage = message;
            }
        }
    }
}
