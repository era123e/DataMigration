// See https://aka.ms/new-console-template for more information
using DataMigration;
using DataMigration.DAL;
using DataMigration.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.Configuration;

string connectionString =
            ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

var options = new DbContextOptionsBuilder<AppDbContext>()
    .UseSqlServer(connectionString)
    .Options;

using var context = new AppDbContext(options);


ExcelOperationServices operations = new ExcelOperationServices(context);

await operations.InportClients();
await operations.InportTechnicians();
await operations.InportWorkOrders();



//Console.WriteLine("Connection loaded successfully:");
//Console.WriteLine(connectionString);



