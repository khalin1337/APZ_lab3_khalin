using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Xml;
using System.Runtime.Serialization.Json;

namespace Khalin_Kypcova_612pst.Classes
{
    public static class Serializacion
    {
        public static void SerialiazeToJson<T>(ref T inObject, string inFileName)
        {
            File.Delete(inFileName);
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T));
            using (FileStream stream1 = new FileStream(inFileName, FileMode.OpenOrCreate))
            {
                ser.WriteObject(stream1, inObject);
                stream1.Close();
            }
        }

        public static void DeserializationFromJson<T>(ref T inObject, string inFileName)
        {
            try
            {
                DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings
                {
                    KnownTypes = new[] { typeof(Admin), typeof(Guest) } // Добавляем Admin как известный тип
                };
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(T), settings);
                using (FileStream stream1 = new FileStream(inFileName, FileMode.Open))
                {
                    inObject = (T)ser.ReadObject(stream1);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Deserialization Error: " + ex.Message);
                throw;
            }
        }
    }
}
