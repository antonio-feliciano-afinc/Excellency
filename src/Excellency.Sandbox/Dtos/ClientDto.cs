using Excellency.Attributes;

namespace Excellency.Sandbox.Dtos
{
    [SheetSource("Clientes")]
    public class ClientDto
    {
        [CellSource("Nome")]
        public string Name  { get; set; }
        
        [CellSource("sobrenome")]
        public string LastName  { get; set; }
        
        [CellSource("telefone")]
        public string Phone  { get; set; }

        [CellSource("cpf")]
        public string CPF { get; set; }

        public string NotMapped { get; set; }
    }
}
