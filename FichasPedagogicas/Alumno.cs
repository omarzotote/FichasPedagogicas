using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FichasPedagogicas
{
    class Alumno
    {
        public String 
            name, 
            firstname,
            lastname;
        public string getName()
        {
            return name + " " + firstname + " " + lastname;
        }
    }
}
