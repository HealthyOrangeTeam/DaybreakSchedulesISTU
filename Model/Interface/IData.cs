using Model.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model.Interface
{
    public interface IData
    {
        AppContext2 _db { get; set; }
        void Add(Week item);
        void Update(Week item);
        void Remove(Week item);
        void Add(Group item);
        void Update(Group item);
        void Remove(Group item);

    }
}
