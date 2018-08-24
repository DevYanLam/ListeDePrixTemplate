using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListeDePrixNovago.Utility
{
    class ListItem
    {
        private string listType;
        private string id;
        private string description;
        private List<Price> price;

        public string ListType { get => listType; set => listType = value; }
        public string Id { get => id; set => id = value; }
        public string Description { get => description; set => description = value; }
        public List<Price> Price { get => price; set => price = value; }
        
    }
}
