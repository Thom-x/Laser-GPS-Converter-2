using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Yaml.Serialization;

namespace Laser_GPS_Converter_2.Config
{
    class ConfigYaml
    {
        private Properties _properties = new Properties();
        private YamlSerializer _serializer = new YamlSerializer();

        public Properties Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }

        private ConfigYaml()
        {
            try
            {
                _properties = (Properties)_serializer.DeserializeFromFile("properties.yml", typeof(Properties))[0];
                if (_properties.DbPath == null || _properties.ExportPath == null)
                {
                    _properties = new Properties();
                    _instance.Update();
                }

                if (!System.IO.File.Exists(_properties.DbPath))
                {
                    _properties.DbPath = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
                }
                if (!System.IO.Directory.Exists(_properties.ExportPath))
                {
                    _properties.ExportPath = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
                }

            }
            catch (Exception)
            {
                _serializer.SerializeToFile("properties.yml", _properties);
            }
        }

        public void Update()
        {
            _serializer.SerializeToFile("properties.yml", _properties);
        }

        private static ConfigYaml _instance;
        public static ConfigYaml Instance
        {
            get 
            {
                if (_instance == null)
                {
                    _instance = new ConfigYaml();
                }
                return _instance;
            }
        }

    }
}
