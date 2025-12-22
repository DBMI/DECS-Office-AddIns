using Bogus;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.ApplicationServices;
using System;


namespace DECS_Excel_Add_Ins
{
    public class AddressModel
    {
        public string Street { get; set; }
        public string StreetNumber {  get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
    }
    public class Patient
    {
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public DateTime DateOfBirth { get; set; }
    }

    internal class GenerateFakeRecords
    {
        private Faker faker;
        private Worksheet worksheet;
        private const int minAge = 18;
        private const int maxAge = 80;
        private DateTime now = DateTime.Now;
        private DateTime maxDateLimit;
        private DateTime minDateLimit;
        private Range target;

        public GenerateFakeRecords(int numRecords)
        {
            faker = new Faker("en_US");
            worksheet = Utilities.CreateNewNamedSheet("Fake Data");

            target = (Range)worksheet.Cells[1, 1];
            BuildHeader();
            
            maxDateLimit = now.AddYears(-minAge);
            minDateLimit = now.AddYears(-maxAge);

            for (int i = 1; i <= numRecords; i++)
            {
                GenerateFakeRow();
            }
        }

        internal void BuildHeader()
        {
            target.Value = "Name";
            target.Offset[0, 1].Value = "Address";
            target.Offset[0, 2].Value = "City";
            target.Offset[0, 3].Value = "State";
            target.Offset[0, 4].Value = "Zip";
            target.Offset[0, 5].Value = "DOB";
            target.Offset[0, 6].Value = "Email";
            target = target.Offset[1, 0];
        }

        internal void GenerateFakeRow()
        {
            var patientFaker = new Faker<Patient>()
                .RuleFor(u => u.FirstName, f => f.Name.FirstName())
                .RuleFor(u => u.LastName, f => f.Name.LastName())
                .RuleFor(u => u.MiddleName, f => f.Name.FirstName())
                .RuleFor(u => u.Email, (f, u) => f.Internet.Email(u.FirstName, u.LastName))
                .RuleFor(u => u.DateOfBirth, f =>
                {
                    // Generate a random date between minDateLimit and maxDateLimit
                    return f.Date.Between(minDateLimit, maxDateLimit);
                });

            var patient = patientFaker.Generate();
            target.Value = patient.LastName + ", " + patient.FirstName + " " + patient.MiddleName;

            var fakeAddresses = new Faker<AddressModel>()
                .RuleFor(a => a.Street, f => f.Address.StreetName())
                .RuleFor(a => a.StreetNumber, f => f.Address.BuildingNumber())
                .RuleFor(a => a.City, f => f.Address.City())
                .RuleFor(a => a.State, f => f.Address.State())
                .RuleFor(a => a.ZipCode, f => f.Address.ZipCode()); // Ensures a zip code is generated
            
            var address = fakeAddresses.Generate();
            target.Offset[0, 1].Value = address.StreetNumber + " " + address.Street;
            target.Offset[0, 2].Value = address.City;
            target.Offset[0, 3].Value = address.State;
            target.Offset[0, 4].Value = address.ZipCode;
            target.Offset[0, 5].Value = patient.DateOfBirth.Date;
            target.Offset[0, 6].Value = patient.Email;
            target = target.Offset[1, 0];
        }
    }
}
