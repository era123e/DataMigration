using DataMigration.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMigration.DAL
{
    public class WorkOrderService
    {
        private readonly AppDbContext _context;
        public WorkOrderService(AppDbContext context)
        {
            _context = context;
        }

        public async Task AddWorkOrderAsync(WorkOrder workOrder)
        {
            _context.WorkOrders.Add(workOrder);
            await _context.SaveChangesAsync();
        }

        public async Task AddAllWorkOrdersAsync(List<WorkOrder> workOrders)
        {
            _context.WorkOrders.AddRange(workOrders);
            await _context.SaveChangesAsync();
        }
    }
}
