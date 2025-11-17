using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMigration.Models
{
    public class WorkOrder
    {
        public int Id { get; set; }

        [Required]
        [ForeignKey(nameof(Technician))]
        public int TechnicianId { get; set; }

        [Required]
        [ForeignKey(nameof(Client))]
        public int ClientId { get; set; }

        public string Information { get; set; }

        public  DateOnly Date { get; set; }

        public decimal Total { get; set; }
        public string RawInformation { get; set; }

    }
}
