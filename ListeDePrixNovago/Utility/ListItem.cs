using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListeDePrixNovago.Utility
{
    class ListItem
    {
        private string id;
        private string description;
        private double price;

        public string Id { get => id; set => id = value; }
        public string Description { get => description; set => description = value; }
        public double Price { get => price; set => price = value; }
    }
}
