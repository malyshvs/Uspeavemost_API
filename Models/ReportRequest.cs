namespace Uspevaemost_API.Models
{
    public class ReportRequest
    {
        public List<string> Kurs { get; set; }
        public List<string> Urovni { get; set; }
        public List<string> FormyObucheniya { get; set; }
        public List<string> Goda { get; set; }
        public List<string> Semestry { get; set; }
        public string name { get; set; }
    }

}
