using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ListeDePrixNovago.PDFTemplate
{
    class SaveXml
    {
        public static void SaveData(object obj, string fileName)
        {
            XmlSerializer sr = new XmlSerializer(obj.GetType());
            TextWriter writer = new StreamWriter(fileName);
            sr.Serialize(writer, obj);
            writer.Close();
        }

        public static PriceListConfig GetData(string fileName)
        { 
            XmlSerializer reader = new XmlSerializer(typeof(PriceListConfig));
            StreamReader file = new StreamReader(fileName);
            PriceListConfig overview = (PriceListConfig)reader.Deserialize(file);
            file.Close();
            return overview;
        }
    }
}
