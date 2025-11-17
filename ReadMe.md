# Data Migration

## Si ta konfigurojme

1. Klonojme repozitonin 
2. Vendosim connectionString per databazen tek App.config dhe tek AppDbContextFactory.cs
3. Ekzekutojme komanden:  **dotnet ef database update**
4. Bejme run projektin nga Visual Studio

## Arkitektura e perdorur

Projekti eshte realizuar me .NET Core. Databaza destinacion per migrim e te dhenave eshte SQL Server. Per aksesimin e te dhenave eshte perdorur Entity Framework Core.
Te dhenat inportohen nga Excel me ane te librarise EPPlus. 
Projekti eshte Console application i cili krijon nje executable file ku kemi mundesine te vendosim ne Task Schedule qe ofron vete Windows.
Path-i per inportimin e te dhenave nga Excel eshte ne folderin InputData, ne te njejten direktori me executable file. 
Emrat e fileve Excel jane statik "Finance-Client.xlsx" dhe "Operations - Work Orders.xlsx". 
Per organizimin e marrjes se te dhenave kemi perdorur services architecture. 

## Gabime te njohura
- Programi duhet te ekzekutohet vetem nje here per cdo file Excel qe do inportohet. 
- Importimi kur data eshte e formateve te ndryshme nga: ["d/M/yyyy", "dd/MM/yyyy", "d/MM/yyyy", "dd/M/yyyy"] nuk eshte i mundur.
- Kolona "RawInformation" tek tabela WorkOrders duhet fshire pasi ajo eshte perdorur per te kontrolluar te dhenat e importuara.
- ExcelOperationServices.cs ka nevoje per refactoring.
- Funksioni **CleanInfrmation()** nuk eshte me optimali i mundshem por arrin nje perqindje te kenaqshme te rezultatit te deshiruar.
- Loget qe dalin ne CSV jane nga nje per cdo rekord te WorkOrders. Ne rast se ka me shume gabime ato dalin ne baze te prioritetit, pra nese rregulohet gabimi per rekord ai mund te kete nje gabim tjeter me prioritet me te ulet.
- Klienti gjendet kur 
  - kane emer dhe mbiemer te sakte
  - kur kane emrin e sakte dhe mbiemrin perafert
  - kur kane mbiemrin e sakte dhe emrin perafert
  - kur jane pjese e emrave dhe mbiemrave te Clients (pra kur kane karaktere mangut ne fillim ose ne fund te emrit dhe mbiemrit)
- klienti nuk gjendet kur ka gabime si ne emer dhe ne mbiemer.  
