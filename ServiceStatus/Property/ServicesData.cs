namespace ServiceStatus.Property
{
    public class ServicesData
    {
        public string Service { get; set; }
        public string MachineIp { get; set; }
        public string ServiceName { get; set; }
        public bool ServiceFlag { get; set; }
        public bool IsServiceCheckSuccess { get; set; }
        public string ExpectedServiceStatus { get; set; }
        public string CurrentServiceStatus { get; set; }
        public bool IsServiceSkipped { get; set; }
        public bool IsStartStopSuccess { get; set; }
        public string StartStopMessage { get; set; }

    }
}
