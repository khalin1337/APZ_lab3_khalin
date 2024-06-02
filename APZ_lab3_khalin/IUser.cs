using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Khalin_Kypcova_612pst.Classes
{

    public interface IUser
    {
        int Id { get; set; }
        string Name { get; set; }
        string Password { get; set; }
        string Email { get; set; }
        string Phone { get; set; }
    }
}
