using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model.Models.Entities
{
    public class Week
    {
        public int Id { get; set; }
        public string PathToImg { get; set; }
        public DateTime DateTime { get; set; }
        public Group Group { get; set; }
        public int GroupId { get; set; }

    }
}
