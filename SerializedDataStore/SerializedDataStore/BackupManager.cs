using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SerializedDataStore
{
    class BackupManager
    {
        const string fileName = @"./addressbook.xml";

        public void Save(AddressBook addressBook)
        {
            using (var sw = new StreamWriter(fileName, false, new System.Text.UTF8Encoding(false)))
            {
                var serializer = new XmlSerializer(typeof(AddressBook));
                serializer.Serialize(sw, addressBook);
            }
        }

        public void Load(ref AddressBook addressBook)
        {
            try
            {
                using (var sr = new StreamReader(fileName, new UTF8Encoding(false)))
                {
                    var serializer = new XmlSerializer(typeof(AddressBook));
                    addressBook = serializer.Deserialize(sr) as AddressBook;
                }
            }
            catch(FileNotFoundException e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
