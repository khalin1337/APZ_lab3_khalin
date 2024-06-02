using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Security.Cryptography;
using System.Runtime.Serialization;

namespace Khalin_Kypcova_612pst.Classes
{
    [DataContract]
    [KnownType(typeof(Admin))]
    public class Admin : IUser
    {
        [DataMember]
        public int Id { get; set; }
        [DataMember]
        public string Email { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string Password { get; set; }
        [DataMember]
        public string Phone { get; set; }
        public Admin() { }

        public Admin(int _id, string _email, string _name, string _password, string _phone)
        {
            Id = _id;
            Email = _email;
            Name = _name;
            Password = _password;
            Phone = _phone;
        }
    }
}
