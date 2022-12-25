using Model.Interface;
using Model.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model.Models.Data
{
    public class Data : IData
    {
        public AppContext2 _db { get; set; }

        public Data()
        {
            _db = new AppContext2();
        }

        public void Add(Week item)
        {

            if (_db.Weeks.Where(q => q.PathToImg == item.PathToImg).FirstOrDefault() != null)
            {
                Update(item);
            }
            else
            {
                _db.Weeks.Add(item);

            }


            _db.SaveChanges();
        }

        public void Add(Group item)
        {
            if (!_db.Groups.Where(g => g.Name == item.Name).Any())
            {
                _db.Groups.Add(item);
            }
            _db.SaveChanges();

        }


        public void Remove(Week item)
        {
            throw new NotImplementedException();
        }

        public void Remove(Group item)
        {
            throw new NotImplementedException();
        }


        public void Update(Week item)
        {
            var dbItem = _db.Weeks.Where(q => q.PathToImg == item.PathToImg).FirstOrDefault();
            if (dbItem != null)
            {
                dbItem = item;
            }

            _db.SaveChanges();

        }

        public void Update(Group item)
        {
            throw new NotImplementedException();
        }
    }
}
