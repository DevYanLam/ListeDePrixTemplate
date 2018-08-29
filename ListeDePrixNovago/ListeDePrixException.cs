using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListeDePrixNovago
{
    class ListeDePrixException: Exception
    {
        public ListeDePrixException()
        {

        }
        public ListeDePrixException(string message) : base(message)
        {

        }
    }
}
