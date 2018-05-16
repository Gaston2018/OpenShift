using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace app.Models
{
    public class MassiveMailAppSetting
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public string User { get; set; }
        public string Password { get; set; }
        public string FilesPath { get; set; }
        public string LogFilesPath { get; set; }
        public string LogFileName { get; set; }
        public string PdfsFilesPath { get; set; }
    }
}
