using Bogus;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;


namespace DECS_Excel_Add_Ins
{
    public class AddressModel
    {
        public string Street { get; set; }
        public string StreetNumber {  get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string StateAbbreviation { get; set; }
        public string ZipCode { get; set; }
    }
    public class Patient
    {
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string MRN { get; set; }
    }

    internal class GenerateFakeRecords
    {
        private Worksheet worksheet;
        private const int minAge = 18;
        private const int maxAge = 80;
        private DateTime now = DateTime.Now;
        private DateTime maxDateLimit;
        private DateTime minDateLimit;
        private Range target;
        private Dictionary<string, string> states;

        public GenerateFakeRecords(int numRecords)
        {
            worksheet = Utilities.CreateNewNamedSheet("Fake Data");

            target = (Range)worksheet.Cells[1, 1];
            BuildHeader();
            BuildStateDictionary();
            
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
            target.Offset[0, 4].Value = "State Abbreviation";
            target.Offset[0, 5].Value = "Zip";
            target.Offset[0, 6].Value = "DOB";
            target.Offset[0, 7].Value = "Email";
            target.Offset[0, 8].Value = "MRN";
            target = target.Offset[1, 0];
        }

        // Source - https://stackoverflow.com/a/5719953
        private void BuildStateDictionary()
        {
            states = new Dictionary<string, string>();
            states.Add("AL", "Alabama");
            states.Add("AK", "Alaska");
            states.Add("AZ", "Arizona");
            states.Add("AR", "Arkansas");
            states.Add("CA", "California");
            states.Add("CO", "Colorado");
            states.Add("CT", "Connecticut");
            states.Add("DE", "Delaware");
            states.Add("DC", "District of Columbia");
            states.Add("FL", "Florida");
            states.Add("GA", "Georgia");
            states.Add("HI", "Hawaii");
            states.Add("ID", "Idaho");
            states.Add("IL", "Illinois");
            states.Add("IN", "Indiana");
            states.Add("IA", "Iowa");
            states.Add("KS", "Kansas");
            states.Add("KY", "Kentucky");
            states.Add("LA", "Louisiana");
            states.Add("ME", "Maine");
            states.Add("MD", "Maryland");
            states.Add("MA", "Massachusetts");
            states.Add("MI", "Michigan");
            states.Add("MN", "Minnesota");
            states.Add("MS", "Mississippi");
            states.Add("MO", "Missouri");
            states.Add("MT", "Montana");
            states.Add("NE", "Nebraska");
            states.Add("NV", "Nevada");
            states.Add("NH", "New Hampshire");
            states.Add("NJ", "New Jersey");
            states.Add("NM", "New Mexico");
            states.Add("NY", "New York");
            states.Add("NC", "North Carolina");
            states.Add("ND", "North Dakota");
            states.Add("OH", "Ohio");
            states.Add("OK", "Oklahoma");
            states.Add("OR", "Oregon");
            states.Add("PA", "Pennsylvania");
            states.Add("RI", "Rhode Island");
            states.Add("SC", "South Carolina");
            states.Add("SD", "South Dakota");
            states.Add("TN", "Tennessee");
            states.Add("TX", "Texas");
            states.Add("UT", "Utah");
            states.Add("VT", "Vermont");
            states.Add("VA", "Virginia");
            states.Add("WA", "Washington");
            states.Add("WV", "West Virginia");
            states.Add("WI", "Wisconsin");
            states.Add("WY", "Wyoming");
        }

        internal void GenerateFakeRow()
        {
            var patientFaker = new Faker<Patient>("en_US")
                .RuleFor(u => u.FirstName, f => f.Name.FirstName())
                .RuleFor(u => u.LastName, f => f.Name.LastName())
                .RuleFor(u => u.MiddleName, f => f.Name.FirstName())
                .RuleFor(u => u.Email, (f, u) => f.Internet.Email(u.FirstName, u.LastName))
                .RuleFor(u => u.MRN, f => f.Random.ReplaceNumbers("########"))
                .RuleFor(u => u.DateOfBirth, f =>
                {
                    // Generate a random date between minDateLimit and maxDateLimit
                    return f.Date.Between(minDateLimit, maxDateLimit);
                });

            var patient = patientFaker.Generate();
            target.Value = patient.LastName + ", " + patient.FirstName + " " + patient.MiddleName;

            var fakeAddresses = new Faker<AddressModel>("en_US")
                .RuleFor(a => a.Street, f => f.Address.StreetName())
                .RuleFor(a => a.StreetNumber, f => f.Address.BuildingNumber())
                .RuleFor(a => a.City, f => f.Address.City())
                .RuleFor(a => a.StateAbbreviation, f => f.Address.StateAbbr())
                .RuleFor(a => a.ZipCode, f => f.Address.ZipCode()); // Ensures a zip code is generated
            
            var address = fakeAddresses.Generate();
            target.Offset[0, 1].Value = address.StreetNumber + " " + address.Street;
            target.Offset[0, 2].Value = address.City;
            target.Offset[0, 3].Value = states[address.StateAbbreviation];
            target.Offset[0, 4].Value = address.StateAbbreviation;
            target.Offset[0, 5].Value = address.ZipCode;
            target.Offset[0, 6].Value = patient.DateOfBirth.Date;
            target.Offset[0, 7].Value = patient.Email;
            target.Offset[0, 8].Value = patient.MRN;
            target = target.Offset[1, 0];
        }
    }
}
