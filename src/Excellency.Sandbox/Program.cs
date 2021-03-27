using Excellency.Enums;
using Excellency.Sandbox.Dtos;
using Excellency.StreamExtensions;
using System.Collections.Generic;
using System.IO;

namespace Excellency.Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            IEnumerable<ClientDto> clients;
            using (FileStream excelstream = File.Open(@"..\..\..\Assets\sampleDataClient.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                clients = excelstream.DeserializeFromExcel<ClientDto>(ExcelFileExtension.Xlsx);
            }

            IEnumerable<WorkDto> works;
            using (FileStream excelstream = File.Open(@"..\..\..\Assets\sampleDataWork.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                works = excelstream.DeserializeFromExcel<WorkDto>(ExcelFileExtension.Xlsx);
            }
        }
    }
}
