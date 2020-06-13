using System.Collections.Generic;

namespace ServiceStatus.Property
{
    public class MachineData
    {
        public string MachineName { get; set; }
        public string MachineIp { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsMachineSkipped { get; set; }
        public bool isSuccess { get; set; }
        public List<ServicesData> services { get; set; }
    }
}
