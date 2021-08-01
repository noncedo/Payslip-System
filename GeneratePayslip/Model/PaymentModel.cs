namespace GeneratePayslip.Model
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class PaymentModel : DbContext
    {
        public PaymentModel()
            : base("name=PaymentModel")
        {
        }

        public virtual DbSet<Employee> Employees { get; set; }
        public virtual DbSet<Payslip> Payslips { get; set; }
        public virtual DbSet<sysdiagram> sysdiagrams { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Payslip>()
                .Property(e => e.BasicSalary)
                .HasPrecision(18, 0);

            modelBuilder.Entity<Payslip>()
                .Property(e => e.PAYE)
                .HasPrecision(18, 0);

            modelBuilder.Entity<Payslip>()
                .Property(e => e.UIF)
                .HasPrecision(18, 0);
        }
    }
}
