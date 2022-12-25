using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model.Models.Entities
{
    public class Group
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public List<Week> Weeks { get; set; }
        public string PathToDocument { get; set; }
    }
}
