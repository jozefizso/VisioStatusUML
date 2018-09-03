namespace VisioStatusUML.Dominio
{
    public class StatusSinistro
    {
        public string CodigoTransacao { get; set; }
        public string NomeTransacao { get; set; }

        public string CodigoStatusAtual { get; set; }
        public string NomeStatusAtual { get; set; }

        public string CodigoStatusSeguinte { get; set; }
        public string NomeStatusSeguinte { get; set; }
    }
}
