namespace NPOItest.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitNPOIModel : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Account",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Username = c.String(nullable: false, maxLength: 255),
                        Password = c.String(nullable: false, maxLength: 256),
                        Name = c.String(nullable: false, maxLength: 256),
                        Email = c.String(nullable: false, maxLength: 256),
                        Sex = c.String(nullable: false, maxLength: 256),
                        Company = c.String(nullable: false, maxLength: 256),
                        Position = c.String(nullable: false, maxLength: 256),
                        Phone = c.String(nullable: false, maxLength: 256),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Account");
        }
    }
}
