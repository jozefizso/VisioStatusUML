using System.Collections.Generic;

namespace VisioStatusUML.Dominio
{
    public class StatusSinistro
    {
        public string CodigoTransacao { get; set; }
        public string NomeTransacao { get; set; }

        public string CodigoStatusAnterior { get; set; }
        public string NomeStatusAnterior { get; set; }

        public string CodigoStatusSeguinte { get; set; }
        public string NomeStatusSeguinte { get; set; }


        public List<StatusSinistro> ListaStatusAnterior { get; set; }
        public List<StatusSinistro> ListaStatusSeguintes { get; set; }
    }
}
