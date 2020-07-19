using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SerializedDataStore
{
    public class PersonalData
    {
        public string Name { set; get; }
        public string Address { set; get; }

        public PersonalData()
        {
            // Not Done
        }

        public PersonalData(string name, string address)
        {
            Name = name;
            Address = address;
        }
    }

    public class AddressBook
    {
        ObservableCollection<PersonalData> addressList = new ObservableCollection<PersonalData>();
        public ObservableCollection<PersonalData> AddressList
        {
            get
            {
                return addressList;
            }
            set
            {
                addressList = value;
            }
        }

        public AddressBook()
        {
            // Not done
        }

        public void Add(PersonalData data)
        {
            addressList.Add(data);
        }
    }
}
