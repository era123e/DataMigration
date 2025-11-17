using DataMigration.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMigration.DAL
{
    public class TechnicianServices
    {

        private readonly AppDbContext _context;

        public TechnicianServices(AppDbContext context)
        {
            _context = context;
        }

        public async Task<List<Technician>> GetAllTechnicians()
        {
            return await _context.Technicians.ToListAsync();
        }

        public async Task AddTechnicianAsync(Technician technician)
        {
            var exists = await _context.Technicians
            .AnyAsync(c => c.FirstName == technician.FirstName && c.LastName == technician.LastName);

            if (exists)
                return;

            _context.Technicians.Add(technician);
            await _context.SaveChangesAsync();
        }

        public async Task AddAllTechniciasnAsync(List<Technician> technicians)
        {
            _context.Technicians.AddRange(technicians);
            await _context.SaveChangesAsync();
        }
    }
}
