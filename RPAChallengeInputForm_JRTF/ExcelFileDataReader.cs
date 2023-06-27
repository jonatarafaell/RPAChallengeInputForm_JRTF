using IronXL;
using RPAChallengeInputForm_JRTF.Models;

namespace RPAChallengeInputForm_JRTF
{
    public class ExcelFileDataReader
    {
        private WorkBook _excelFile;
        private WorkSheet _excelSpreadsheet;
        private List<EmployeeModel> _employeeList;

        public WorkBook ExcelFile
        {
            get { return _excelFile; }
            set { _excelFile = value; }
        }

        public WorkSheet ExcelSpreadsheet
        {
            get { return _excelSpreadsheet; }
            set { _excelSpreadsheet = value; }
        }

        public List<EmployeeModel> EmployeeList
        {
            get { return _employeeList; }
            set { _employeeList = value; }
        }

        /*  O método recebe como parâmentro o caminho da planilha do Excel que já está presente no projeto.
         *  Em seguida, o método realiza a leitura dos valores de todas as células de uma mesma linha, pulando a primeira linha (cabeçalho),
         *  e armazena em um objeto os valores lidos. Por fim, cada objeto é armazenado em uma lista e esta é retornada.
         */
        public List<EmployeeModel> ReadEmployeeDataFromSpreadsheet(string spreadsheetPath)
        {
            ExcelFile = WorkBook.Load(spreadsheetPath);
            ExcelSpreadsheet = ExcelFile.WorkSheets.First();

            EmployeeList = new List<EmployeeModel>();
            var employee = new EmployeeModel();

            foreach (var cell in ExcelSpreadsheet["A2:G11"])
            {
                switch (cell.ColumnIndex)
                {
                    case 0:
                        employee.FirstName = cell.StringValue;
                        break;

                    case 1:
                        employee.LastName = cell.StringValue;
                        break;

                    case 2:
                        employee.CompanyName = cell.StringValue;
                        break;
                    case 3:
                        employee.RoleInCompany = cell.StringValue;
                        break;
                    case 4:
                        employee.Address = cell.StringValue;
                        break;
                    case 5:
                        employee.Email = cell.StringValue;
                        break;
                    case 6:
                        employee.PhoneNumber = cell.StringValue;
                        EmployeeList.Add(employee);
                        employee = new EmployeeModel();
                        break;

                    default:
                        break;
                }
            }

            return EmployeeList;
        }
    }
}