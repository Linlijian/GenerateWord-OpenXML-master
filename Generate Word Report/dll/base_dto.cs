using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Generate_Word_Report.dll
{
    public class base_dto
    {
        private ErrorResults dtoErrorResults = new ErrorResults();

        public ErrorResults ErrorResults
        {
            get { return dtoErrorResults; }
            set { dtoErrorResults = value; }
        }
    }

    public class ErrorResults
    {
        public string error_message { get; set; }
        public int error_code { get; set; }
    }
}
