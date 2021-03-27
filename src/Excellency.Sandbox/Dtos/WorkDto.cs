using Excellency.Attributes;
using System;

namespace Excellency.Sandbox.Dtos
{
    [SheetSource("WOs")]
    public class WorkDto
    {
        [CellSource("WO")]
        public string WO { get; set; }

        [CellSource("District")]
        public string Distrito { get; set; }

        [CellSource("LeadTech")]
        public string Lider { get; set; }

        [CellSource("Service")]
        public string Servico { get; set; }

        [CellSource("ReqDate")]
        public DateTime DataReq { get; set; }

        [CellSource("WorkDate")]
        public DateTime DataTrabalho { get; set; }        
        
        [CellSource("TotalCost")]
        public decimal Custo { get; set; }
        

    }
}
