using IronXL;
using RPA.Models;

namespace RPA
{
    public class ExcelDataReader
    {
        // classe da biblioteca IronXL que representa um arquivo Excel inteiro. 
        private WorkBook _excelFile;
        //  classe da biblioteca IronXL que representa uma aba específica dentro do arquivo Excel.
        private WorkSheet _excelSpreadSheet;
        private List<PersonModel> _personList;

        public WorkBook ExcelFile { get { return _excelFile; } set { _excelFile = value; } }
        public WorkSheet ExcelSpreadSheet { get { return _excelSpreadSheet; } set { _excelSpreadSheet = value; } }
        public List<PersonModel> PersonList { get { return _personList; }  set { _personList = value; } }

        public List<PersonModel> ReadingSpreadsheet(string excelPath)
        {
            ExcelFile = WorkBook.Load(excelPath);

            ExcelSpreadSheet = ExcelFile.WorkSheets.First();

            PersonList = new List<PersonModel>();           

            foreach (var row in ExcelSpreadSheet.Rows.Skip(1).Take(10))
            {
                var person = new PersonModel
                {
                    FirstName = row.Columns[0].StringValue,
                    LastName = row.Columns[1].StringValue,
                    CompanyName = row.Columns[2].StringValue,
                    RoleInCompany = row.Columns[3].StringValue,
                    Address = row.Columns[4].StringValue,
                    Email = row.Columns[5].StringValue,
                    PhoneNumber = row.Columns[6].StringValue
                };

                PersonList.Add(person);
            }            
            return PersonList;
        }
    }

}
