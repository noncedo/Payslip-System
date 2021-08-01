using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneratePayslip.Model
{
    class EmployeePayslipModel
    {

        public virtual Employee EmployeeModel { get; set; }
        public virtual Payslip PayslipModel { get; set; }
    }
}
