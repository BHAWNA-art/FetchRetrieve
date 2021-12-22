using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace FetchRetreive.Models
{
    public class Uploadfile
    {
        [Required]
        public HttpPostedFileBase Excelfile { get; set; }
    }
}