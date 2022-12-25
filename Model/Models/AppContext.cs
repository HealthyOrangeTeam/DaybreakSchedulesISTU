using Microsoft.EntityFrameworkCore;
using Model.Models;
using Model.Models.Entities;

public class AppContext2 : DbContext
{
    public DbSet<Group> Groups { get; set; }
    public DbSet<Week> Weeks { get; set; }

    public AppContext2()
    {
        //Database.EnsureDeleted();
        Database.EnsureCreated();
    }


    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlite("Data Source=helloapp.db");
    }
}