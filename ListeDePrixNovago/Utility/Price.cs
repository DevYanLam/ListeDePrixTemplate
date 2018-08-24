using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListeDePrixNovago.Utility
{
    public class Price
    {
        private string name;
        private double amount;
        private bool isChecked;

        public string Name { get => name; set => name = value; }
        public bool IsChecked { get => isChecked; set => isChecked = value; }
        public double Amount { get => amount; set => amount = value; }
    }
}
