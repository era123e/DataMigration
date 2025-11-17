using DataMigration.Models;
using Microsoft.SqlServer.Server;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static OfficeOpenXml.ExcelErrorValue;

namespace DataMigration.DAL
{
    public class ExcelOperationServices
    {
        const int INSERT_BATCH_LIMIT = 100;


        private readonly ClientService clientService;
        private readonly TechnicianServices technicianService;
        private readonly WorkOrderService workOrderService;
        public ExcelOperationServices(AppDbContext context)
        { 
            clientService = new ClientService(context);
            technicianService = new TechnicianServices(context);
            workOrderService = new WorkOrderService(context);
        }

        public async Task InportClients()
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            string relativeClientsPath = Path.Combine("InputData", "Finance - Clients.xlsx");
            string clientsFilePath = Path.Combine(AppContext.BaseDirectory, relativeClientsPath);

            if (!File.Exists(clientsFilePath))
            {
                Console.WriteLine($"File not found: {clientsFilePath}");
                return;
            }

            using var finance = new ExcelPackage(new FileInfo(clientsFilePath));
            var worksheet = finance.Workbook.Worksheets[0];

            int rowCount = worksheet.Dimension.Rows;

            List<Client> clientList = new List<Client>();

            for (int row = 2; row <= rowCount; row++) // row = 2: skip header
            {
                string fullName = worksheet.Cells[row, 1].Text;

                if (!string.IsNullOrWhiteSpace(fullName))
                {
                    var parts = fullName.Split(' ', 2);
                    string firstName = parts.Length > 0 ? parts[0] : "";
                    string lastName = parts.Length > 1 ? parts[1] : "";

                    Console.WriteLine($"Row {row}: FirstName='{firstName}', LastName='{lastName}'");

                    var newClient = new Client
                    {
                        FirstName = firstName,
                        LastName = lastName
                    };

                    clientList.Add(newClient);

                    // Menyra I: e avashte pasi ruan ne databaze per cdo rresht
                    //await clientService.AddClientAsync(newClient);

                    // Menyra II: me e shpejte pasi ne databaze i ruajme ne grupe
                    // por duhet te kemi kujdes qe mos i ruajme te gjitha njeheresh pasi mund te konsumojme resource te serverit te aplikacionit
                    // keshtu qe i ruajme me pjese sipas limitit te konfiguruar.

                    if (clientList.Count == INSERT_BATCH_LIMIT || row == rowCount)  // row == rowCount: bejme kujdes qe edhe nese nuk arrihet limiti, pjesa e fundit te mos ngeli pa u bere import
                    {
                        await clientService.AddAllClientsAsync(clientList);
                        clientList.Clear();
                    }
                }
            }

            Console.WriteLine("Clients imported in database!");
        }

        public async Task InportTechnicians()
        {

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            string relativeOperationsPath = Path.Combine("InputData", "Operations - Work Orders.xlsx");
            string operationsFilePath = Path.Combine(AppContext.BaseDirectory, relativeOperationsPath);

            if (!File.Exists(operationsFilePath))
            {
                Console.WriteLine($"File not found: {operationsFilePath}");
                return;
            }

            using var operations = new ExcelPackage(new FileInfo(operationsFilePath));
            var operationsWorksheet = operations.Workbook.Worksheets[0];

            int operationRows = operationsWorksheet.Dimension.Rows;

            List<Technician> dbTechnicians = await technicianService.GetAllTechnicians();
            List<Technician> technicianList = new List<Technician>();

            for (int row = 2; row <= operationRows; row++) // row = 2: skip header
            {
                string fullName = operationsWorksheet.Cells[row, 1].Text;

                if (!string.IsNullOrWhiteSpace(fullName))
                {
                    var parts = fullName.Split(' ', 2);
                    string firstName = parts.Length > 0 ? parts[0] : "";
                    string lastName = parts.Length > 1 ? parts[1] : "";

                    Console.WriteLine($"Row {row}: FirstName='{firstName}', LastName='{lastName}'");

                    var newTechnichian = new Technician
                    {
                        FirstName = firstName,
                        LastName = lastName
                    };

                    //technicianList.Add(newTechnichian);

                    // Menyra I: e avashte pasi ruan ne databaze per cdo rresht
                    //await technicianService.AddTechnicianAsync(newTechnichian);

                    // Menyra II: me e shpejte pasi ne databaze i ruajme ne grupe
                    // por duhet te kemi kujdes qe mos i ruajme te gjitha njeheresh pasi mund te konsumojme resource te serverit te aplikacionit
                    // keshtu qe i ruajme me pjese sipas limitit te konfiguruar.

                    // Meqe kemi shtuar kontroll per tekniket duke marre parasysh qe mund te mos jene te njejte
                    // atehere menyra e dyte nuk hyn ne pune pasi nuk eshte e mundur kontrolli.

                    // Per te ritur performancen mund te perdorim librari si EFCore.BulkExtensions
                    // per qellimet e nje demoje po e realizojme ne menyre te shpejte me cfare ofron vete libraria e Entity Framework Core

                    //if (technicianList.Count == INSERT_BATCH_LIMIT || row == operationRows)  // row == rowCount: bejme kujdes qe edhe nese nuk arrihet limiti, pjesa e fundit te mos ngeli pa u bere import
                    //{
                    //    await technicianService.AddAllTechniciasnAsync(technicianList);
                    //    technicianList.Clear();
                    //}

                    // Menyra III: duke krijuar nje liste unike per ti futur ne databaze
                    // Duke supozuar qe tekniket jane ne nje numer te pranueshem per tu ruajtur ne memorje ne nje liste
                    // dhe pastaj te ruhen ne grup ne memorje persistente (databaze)
                    // mund te bejme mbushjen e listes me teknike ne menyre unike dhe i ruajme te gjithe njekohesisht

                    // TODO: kontroll nga databaza per tekniket e futur me pare nga exel per vitet e kaluara
                    // nese kemi teknike jo ne numer shume te madh, i ngarkojme ne liste per ti perdorur per kontroll,
                    // perndryshe na duhet te bejme kerkim ne databaze

                    if (!dbTechnicians.Any(t => t.FirstName == newTechnichian.FirstName && t.LastName == newTechnichian.LastName)) // te mos ekzistoje ne db
                    {
                        if (!technicianList.Any(t => t.FirstName == newTechnichian.FirstName && t.LastName == newTechnichian.LastName))
                        {
                            technicianList.Add(newTechnichian);
                        }
                    }

                }
            }

            await technicianService.AddAllTechniciasnAsync(technicianList);

            Console.WriteLine("Technicians imported in database!");

        }

        public async Task InportWorkOrders()
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            string relativeOperationsPath = Path.Combine("InputData", "Operations - Work Orders.xlsx");
            string operationsFilePath = Path.Combine(AppContext.BaseDirectory, relativeOperationsPath);

            if(!File.Exists(operationsFilePath)) 
            {
                Console.WriteLine($"File not found: {operationsFilePath}");
                return;
            }
            List<string> scvLog = new List<string>();
            List<WorkOrder> workOrders = new List<WorkOrder>();
            List<Technician> allTechnicians = await technicianService.GetAllTechnicians();
            
            decimal totalValue = 0;
            DateTime workOrderDate;

            using var operations = new ExcelPackage(new FileInfo(operationsFilePath));
            var operationsWorksheet = operations.Workbook.Worksheets[0];

            int operationRows = operationsWorksheet.Dimension.Rows;

            for (int row = 2; row <= operationRows; row++) // row = 2: skip header
            {
                int technicianId = 0;
                int clientId = 0;

                bool isFirstNameDifferent = false;
                bool islastNameDifferent = false;
                bool isDateCorrect = true;
                bool generalError = false;
                bool technicanNotFound = false;
                bool totalNotANumber = false;

                string technicianName = operationsWorksheet.Cells[row, 1].Text;
                string rawInformation = operationsWorksheet.Cells[row, 2].Text;
                string total = operationsWorksheet.Cells[row, 3].Text;

                if (!string.IsNullOrWhiteSpace(technicianName))
                {
                    var parts = technicianName.Split(' ', 2);
                    string firstName = parts.Length > 0 ? parts[0] : "";
                    string lastName = parts.Length > 1 ? parts[1] : "";

                    //Console.WriteLine($"Row {row}: FirstName='{firstName}', LastName='{lastName}'");

                    var technician = allTechnicians.Where(t => t.FirstName == firstName && t.LastName == lastName).FirstOrDefault();
                    if (technician != null)
                    {
                        technicianId = technician.Id;
                    }
                    else
                    {
                        technicanNotFound = true;
                    }

                }

                if (decimal.TryParse(total, out totalValue))
                {
                    Console.WriteLine(totalValue);
                }
                else
                {
                    totalNotANumber = true;
                    Console.WriteLine("Invalid decimal");
                }

                var clientFullName = "";

                var nameRegex = new Regex(@"\b\p{Lu}\p{L}+\s\p{Lu}\p{L}+\b");

                clientFullName = nameRegex.Matches(rawInformation).FirstOrDefault().Value;
                Console.WriteLine(clientFullName);
                
                var nameParts = clientFullName.Split(' ', 2);
                var clientName = nameParts.Length > 0 ? nameParts[0] : "";
                var clientSurname = nameParts.Length > 1 ? nameParts[1] : "";

                // per permiresim mund te bejme nje funskion qe merr klientet nga nje liste me emra klientesh
                // keshtu nuk do duhet te bejme query ne databaze per cdo klient
                var client =await clientService.GetClientByFullNameAsync(clientName, clientSurname);

                if (client != null)
                {
                    clientId = client.Id;
                }
                else //TODO: kur nuk ka klientId me emer fiks, gjejme me emer te perafert per te marre parasysh rastin e gabimeve
                {
                    // gjejme klientet me emer ekzakt
                    var clientsByName = await clientService.GetClientByFirstNameAsync(clientName);
                    if (clientsByName != null && clientsByName.Count > 0)
                    {
                        islastNameDifferent = true;

                        var changeDifference = 1000;

                        foreach (var c in clientsByName)
                        {
                            int diff = clientSurname.Zip(c.LastName, (a, b) => a != b).Count(x => x);
                            if (diff < changeDifference)
                            {
                                changeDifference = diff;
                                clientId = c.Id;
                            }

                            // TODO: log in excel qe klienti eshte gjetur me perafersi
                        }
                    }
                    else // gjejme klientet me mbiemer ekzakt
                    {
                        isFirstNameDifferent = true;
                        // gjejme klientet me emer ekzakt
                        var clientsBySurname = await clientService.GetClientByLastNameAsync(clientSurname);
                        if (clientsBySurname != null && clientsBySurname.Count > 0)
                        {
                            var changeDifference = 1000;

                            foreach (var c in clientsBySurname)
                            {
                                int diff = clientName.Zip(c.FirstName, (a, b) => a != b).Count(x => x);
                                if (diff < changeDifference)
                                {
                                    changeDifference = diff;
                                    clientId = c.Id;
                                }

                                // TODO: log in excel qe klienti eshte gjetur me perafersi
                            }

                            // TODO: log in excel qe klienti eshte gjetur me perafersi
                        }
                        else // TODO: Kerkojme me emeer dhe mbiemer me LIKE
                        {
                            isFirstNameDifferent = true;
                            islastNameDifferent = true;

                        }
                    }
                }

                string dateString = "";
                var dateRegex = new Regex(@"\b([0-3]?\d)/([0-1]?\d)/(\d{4})\b");

                dateString = dateRegex.Matches(rawInformation).FirstOrDefault()?.Value;
                Console.WriteLine(dateString);


                if (DateTime.TryParseExact(dateString, ["d/M/yyyy", "dd/MM/yyyy", "d/MM/yyyy", "dd/M/yyyy"], CultureInfo.InvariantCulture, DateTimeStyles.None, out workOrderDate))
                {
                    Console.WriteLine($"{dateString} -> {workOrderDate:yyyy-MM-dd}");
                }
                else
                {
                    isDateCorrect = false;
                    Console.WriteLine($"{dateString} is not a valid date");
                }

                if (isDateCorrect == true && technicanNotFound == false && clientId != 0 && totalNotANumber == false)
                {
                    var workOrder = new WorkOrder()
                    {
                        TechnicianId = technicianId,
                        ClientId = clientId,
                        Total = totalValue,
                        Information = CleanInfrmation(rawInformation, clientFullName, dateString),
                        Date = DateOnly.FromDateTime(workOrderDate),
                        RawInformation = rawInformation
                    };

                    workOrders.Add(workOrder);
                }

                // me e shpejte pasi ne databaze i ruajme ne grupe
                // por duhet te kemi kujdes qe mos i ruajme te gjitha njeheresh pasi mund te konsumojme resource te serverit te aplikacionit
                // keshtu qe i ruajme me pjese sipas limitit te konfiguruar.

                if (workOrders.Count == INSERT_BATCH_LIMIT || row == operationRows)  // row == rowCount: bejme kujdes qe edhe nese nuk arrihet limiti, pjesa e fundit te mos ngeli pa u bere import
                {
                    await workOrderService.AddAllWorkOrdersAsync(workOrders);
                    workOrders.Clear();
                }


                //Shkruajme log ne CSV
                if (totalNotANumber == true)
                {
                    scvLog.Add(row + "," + "Problem me totalin!"); // Error
                }
                else if(technicanNotFound == true)
                {
                    scvLog.Add(row + "," + "Tekniku nuk u gjet!"); // Error
                }
                else if (isDateCorrect == false)
                {
                    scvLog.Add(row + "," + "Formati i dates eshte gabim!"); // Error
                }
                else if (clientId == 0)
                {
                    scvLog.Add(row + "," + "Klienti nuk u gjet!");  // Error
                }
                else if(generalError == true)
                {
                    scvLog.Add(row + "," + "Klienti nuk u gjet!");  // Error
                }
                else if (isFirstNameDifferent == true || islastNameDifferent == true)
                {
                    scvLog.Add(row + "," + "Klienti u gjet me perafersi! Rekordi eshte futur ne databaze por do kontrolluar manualisht."); // Warning
                }
                else
                {
                    scvLog.Add(row + "," + "Importuar me sukses!");  // Sukses
                }

            }

            File.WriteAllLines("errorLog.scv", scvLog);


        }

        private string CleanInfrmation(string rawString, string fullName, string dateString)
        {
            string returString = rawString;

            // remove date
            if (dateString != null)
            {
                returString = returString.Replace(" me daten " + dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();
                returString = returString.Replace("me daten " + dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();
                returString = returString.Replace(" ne daten " + dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();
                returString = returString.Replace("ne daten " + dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();

                returString = returString.Replace(" me " + dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();

                returString = returString.Replace("me " + dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();

                returString = returString.Replace("data " + dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();

                returString = returString.Replace(dateString, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();

                returString = returString = returString.TrimStart('.');
                returString = returString.Trim();
            }

            //remove name
            if(fullName != null)
            {
                returString = returString.Replace(" te " + fullName, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();
                returString = returString.Replace(" per " + fullName, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();

                returString = returString.Replace(fullName, "", StringComparison.OrdinalIgnoreCase);
                returString = returString.Trim();
            }
            returString = returString.Replace(" .", ".");
            returString = returString.Trim(';');
            returString = returString.Trim();
            returString = returString.Trim(',');
            returString = returString.Trim();
            returString = char.ToUpper(returString[0]) + returString.Substring(1);
            
            return returString;
        }
    }
}
