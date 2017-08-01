using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace NPOItest.Models
{
    public class InitDatabase : DropCreateDatabaseAlways<NPOIModel>
    {
        protected override void Seed(NPOItest.Models.NPOIModel context)
        {
            Account acc1 = new Account { Username = "admin", Password = "admin", Name = "Kevin", Sex = "Man", Email = "admin@NPOT.com", Company = "XXX_NO.1", Position = "R&D", Phone = "0000-0000" };
            Account acc2 = new Account { Username = "user", Password = "user", Name = "Durant", Sex = "Female", Email = "user@NPOT.com", Company = "AAA_NO.2", Position = "CEO", Phone = "0000-1111" };

            context.Account.Add(acc1);
            context.Account.Add(acc2);

            context.SaveChanges();
        }
    }
}