using System.ComponentModel.DataAnnotations;
using System.Xml.Linq;

namespace MVC_Music.ViewModels
{
    public class MusicianReportsSummaryVM
    {
        public int ID { get; set; }


        [Display(Name = "Musician")]
        public string FormalName
        {
            get
            {
                return LastName + ", " + FirstName
                    + (string.IsNullOrEmpty(MiddleName) ? "" :
                        (" " + (char?)MiddleName[0] + ".").ToUpper());
            }
        }
        [Display(Name = "First Name")]
        [Required(ErrorMessage = "You cannot leave the first name blank.")]
        [StringLength(30, ErrorMessage = "First name cannot be more than 30 characters long.")]
        public string FirstName { get; set; }

        [Display(Name = "Middle Name")]
        [StringLength(30, ErrorMessage = "Middle name cannot be more than 30 characters long.")]
        public string MiddleName { get; set; }

        [Display(Name = "Last Name")]
        [Required(ErrorMessage = "You cannot leave the last name blank.")]
        [StringLength(50, ErrorMessage = "Last name cannot be more than 50 characters long.")]
        public string LastName { get; set; }


        [Display(Name ="Average Fee Paid")]
        [DataType(DataType.Currency)]
        public double AverageFeePaid { get; set; }


        [Display(Name = "Highest Fee Paid")]
        [DataType(DataType.Currency)]
        public double HighestFeePaid { get; set; }

        [Display(Name = "Lowest Fee Paid")]
        [DataType(DataType.Currency)]
        public double LowestFeePaid { get; set; }


        [Display(Name = "Total Number of Performances")]
        public int NumberofPerformances { get; set; }

    }
}
