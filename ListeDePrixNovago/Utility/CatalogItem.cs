using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListeDePrixNovago.Utility
{
    class CatalogItem
    {
        private string attribut1;
        private string attribut2;
        private string id;
        private string description;
        private string um;
        private List<Price> prices;

        public string Attribut1 { get => attribut1; set => attribut1 = value; }
        public string Attribut2 { get => attribut2; set => attribut2 = value; }
        public string Id { get => id; set => id = value; }
        public string Description { get => description; set => description = value; }
        public string Um { get => um; set => um = value; }
        public List<Price> Prices { get => prices; set => prices = value; }

    }
}
