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
        private double prix1;
        private double prix2;
        private double prix3;
        private double prix4;
        private double prix5;

        public string Attribut1 { get => attribut1; set => attribut1 = value; }
        public string Attribut2 { get => attribut2; set => attribut2 = value; }
        public string Id { get => id; set => id = value; }
        public string Description { get => description; set => description = value; }
        public string Um { get => um; set => um = value; }
        public double Prix1 { get => prix1; set => prix1 = value; }
        public double Prix2 { get => prix2; set => prix2 = value; }
        public double Prix3 { get => prix3; set => prix3 = value; }
        public double Prix4 { get => prix4; set => prix4 = value; }
        public double Prix5 { get => prix5; set => prix5 = value; }
    }
}
