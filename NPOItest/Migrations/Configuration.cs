namespace NPOItest.Migrations
{
    using Models;
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;

    internal sealed class Configuration : DbMigrationsConfiguration<NPOItest.Models.NPOIModel>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }

        protected override void Seed(NPOItest.Models.NPOIModel context)
        {
            //  This method will be called after migrating to the latest version.

            //  You can use the DbSet<T>.AddOrUpdate() helper extension method 
            //  to avoid creating duplicate seed data. E.g.
            //
            //    context.People.AddOrUpdate(
            //      p => p.FullName,
            //      new Person { FullName = "Andrew Peters" },
            //      new Person { FullName = "Brice Lambson" },
            //      new Person { FullName = "Rowan Miller" }
            //    );
            //
            context.Account.AddOrUpdate(x => x.Id,
                new Account { Username = "admin", Password = "admin", Name = "Kevin", Sex = "Man", Email = "admin@NPOT.com", Company = "XXX_NO.1", Position = "R&D", Phone = "0000-0000" },
                new Account { Username = "user", Password = "user", Name = "Durant", Sex = "Female", Email = "user@NPOT.com", Company = "AAA_NO.2", Position = "CEO", Phone = "0000-1111" }
                );
        }
    }
}
