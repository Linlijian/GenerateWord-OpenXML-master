
namespace Pikunword
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
