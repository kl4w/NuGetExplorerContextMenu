using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace NuGetContextMenuHandler
{
    public class ConfigHandler
    {
        private static string optionConfig = String.Format(@"{0}\NuGet\ContextMenu.Config", Environment.GetEnvironmentVariable("LOCALAPPDATA"));

        public IEnumerable<KeyValuePair<string, string>> Sources 
        {
            get
            {
                return GetSources();
            }
        }

        public IEnumerable<KeyValuePair<string, string>> Options 
        {
            get
            {
                return GetOptions();
            }
        }

        public ConfigHandler()
        {
        }

        private IEnumerable<KeyValuePair<string, string>> GetSources()
        {
            var nugetConfig = String.Format(@"{0}\NuGet\NuGet.Config", Environment.GetEnvironmentVariable("APPDATA"));
            
            var allSources = ConfigElements(nugetConfig, "packageSources");
            var disabledSources = ConfigElements(nugetConfig, "disabledPackageSources");
            foreach (var disabledSource in disabledSources)
            {
                var item = allSources.First(x => x.Key == disabledSource.Key);
                allSources.Remove(item);
            }
            return allSources;
        }

        private List<KeyValuePair<string, string>> ConfigElements(string filePath, string tagName)
        {
            List<KeyValuePair<string, string>> elementList = new List<KeyValuePair<string, string>>();

            XmlDocument config = new XmlDocument();
            config.Load(filePath);
            var elements = config.GetElementsByTagName(tagName);
            for (int i = 0; i < elements[0].ChildNodes.Count; i++)
            {
                var source = elements[0].ChildNodes[i].Attributes;
                var key = source.GetNamedItem("key").Value;
                var value = source.GetNamedItem("value").Value;
                KeyValuePair<string, string> item = new KeyValuePair<string, string>(key, value);
                elementList.Add(item);
            }

            return elementList;
        }

        public KeyValuePair<string, string> ConfigElement(string tagName)
        {
            XmlDocument config = new XmlDocument();
            config.Load(optionConfig);
            var elements = config.GetElementsByTagName(tagName);
            var source = elements[0].ChildNodes[0].Attributes;
            var value = source.GetNamedItem("value").Value;
            KeyValuePair<string, string> item = new KeyValuePair<string, string>(tagName, value);
            return item;
        }

        private IEnumerable<KeyValuePair<string, string>> GetOptions()
        {
            if (!File.Exists(optionConfig))
            {
                return SetDefaultOptions(optionConfig);
            }

            List<KeyValuePair<string, string>> options = new List<KeyValuePair<string, string>>();
            options.Add(ConfigElement("useVersionedPackages"));
            options.Add(ConfigElement("getLatest"));
            options.Add(ConfigElement("cleanPackages"));
            return options;
        }

        private static List<KeyValuePair<string, string>> SetDefaultOptions(string path)
        {
            XmlDocument config = new XmlDocument();
            XmlDeclaration dec = config.CreateXmlDeclaration("1.0", "utf-8", null);
            config.AppendChild(dec);
            XmlElement root = config.CreateElement("configuration");
            config.AppendChild(root);

            root.AppendChild(CreateOptionElement(config, "useVersionedPackages", false));
            root.AppendChild(CreateOptionElement(config, "getLatest", true));
            root.AppendChild(CreateOptionElement(config, "cleanPackages", true));

            config.Save(path);

            List<KeyValuePair<string, string>> options = new List<KeyValuePair<string, string>>();
            options.Add(new KeyValuePair<string, string>("useVersionedPackages", "False"));
            options.Add(new KeyValuePair<string, string>("getLatest", "True"));
            options.Add(new KeyValuePair<string, string>("cleanPackages", "True"));
            return options;
        }

        private static XmlElement CreateOptionElement(XmlDocument config, string elementName, bool value)
        {
            XmlElement element = config.CreateElement(elementName);
            XmlElement add = config.CreateElement("add");
            SetValue(add, value);
            element.AppendChild(add);
            return element;
        }

        private static XmlElement SetValue(XmlElement element, bool value)
        {
            element.SetAttribute("key", "enabled");
            element.SetAttribute("value", value.ToString());
            return element;
        }

        public void SwitchOptionValue(int pos)
        {
            XmlDocument config = new XmlDocument();
            config.Load(optionConfig);
            var elements = config.GetElementsByTagName(Options.ElementAt(pos).Key);
            var option = (XmlElement)elements[0].ChildNodes[0];
            SetValue(option, !bool.Parse(option.Attributes.GetNamedItem("value").Value));
            config.Save(optionConfig);
        }
    }
}
