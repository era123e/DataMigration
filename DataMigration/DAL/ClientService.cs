using DataMigration.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMigration.DAL
{
    public class ClientService
    {
        private readonly AppDbContext _context;
        public ClientService(AppDbContext context)
        {
            _context = context;
        }

        public async Task<Client> GetClientByFullNameAsync(string firstName, string lastname)
        {
            return await _context.Clients.Where(c => c.FirstName == firstName && c.LastName == lastname).FirstOrDefaultAsync();            
        }

        public async Task<List<Client>> GetClientByFirstNameAsync(string firstName)
        {
            return await _context.Clients.Where(c => c.FirstName == firstName).ToListAsync();
        }
        public async Task<List<Client>> GetClientByLastNameAsync(string lastName)
        {
            return await _context.Clients.Where(c => c.LastName == lastName).ToListAsync();
        }

        public async Task<List<Client>> GetClientByApproximateNameAsync(string firstName, string lastName)
        {
            return await _context.Clients.Where(c => EF.Functions.Like(c.FirstName, firstName) && EF.Functions.Like(c.LastName, lastName)).ToListAsync();
        }

        public async Task AddClientAsync(Client client)
        {
            _context.Clients.Add(client);
            await _context.SaveChangesAsync();
        }

        public async Task AddAllClientsAsync(List<Client> clients)
        {
            _context.Clients.AddRange(clients);
            await _context.SaveChangesAsync();
        }
    }
}
