namespace GeneratePayslip.Model
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Payslip")]
    public partial class Payslip
    {
        public int PayslipId { get; set; }

        public int? EmployeeId { get; set; }

        public decimal? BasicSalary { get; set; }

        public decimal? PAYE { get; set; }

        public decimal? UIF { get; set; }

        public DateTime? PayDate { get; set; }

        public virtual Employee Employee { get; set; }
    }
}
