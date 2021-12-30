using System;
using System.ComponentModel.DataAnnotations;

namespace AspnetCoreTODO.Models
{
    public class Todo
    {
        [Key]
        public int id { get; set; }
        public string Name { get; set; }
        public DateTime Createdate { get; set; }
  }
}