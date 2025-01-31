﻿namespace ServiceStatus.Property
{
    public class ServicesData
    {
        public string Service { get; set; }
        public string MachineIp { get; set; }
        public string ServiceName { get; set; }
        public bool skipService { get; set; }
        public bool IsServiceCheckSuccess { get; set; }
        public string ExpectedServiceStatus { get; set; }
        public string CurrentServiceStatus { get; set; }
        public bool IsServiceSkipped { get; set; }
        public bool IsStartStopSuccess { get; set; }
        public bool IsStarted { get; set; }
        public bool IsStopped { get; set; }
        public bool IsServiceStatusAsExpected { get; set; }
        public string StartStopMessage { get; set; }

    }
}
