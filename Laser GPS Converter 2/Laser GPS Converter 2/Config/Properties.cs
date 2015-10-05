using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laser_GPS_Converter_2.Config
{
    class Properties
    {
        public Properties()
        {
            string defaultinstalldir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Ultrasport GPS");
            if (System.IO.Directory.Exists(defaultinstalldir))
            {
                _dbPath = defaultinstalldir;
            }
            else
            {
                _dbPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            }
            _exportPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        private string _dbPath;
        private string _exportPath;

        public string DbPath
        {
            get { return _dbPath; }
            set { _dbPath = value; }
        }

        public string ExportPath
        {
            get { return _exportPath; }
            set { _exportPath = value; }
        }
    }
}
