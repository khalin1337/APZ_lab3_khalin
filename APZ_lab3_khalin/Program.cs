using APZ_lab3_khalin;
using Khalin_Kypcova_612pst.Classes;
using System;
using System.Threading;
using System.Threading.Tasks;
namespace Program
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Order> orders = new List<Order>();
            List<IUser> users = new List<IUser>();
            Serializacion.DeserializationFromJson(ref users, "Users.json");
            Serializacion.DeserializationFromJson(ref orders, "Orders.json");
            CreateWordList.CreateList(orders, users);
            CreateExcelList.CreatelList(orders, users);
            Console.WriteLine("qweqweqweqwe");
        }
    }
}