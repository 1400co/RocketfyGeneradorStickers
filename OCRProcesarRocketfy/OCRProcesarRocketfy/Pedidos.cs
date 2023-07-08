using System.Runtime;

namespace OCRProcesarRocketfy
{
    public class Pedidos
    {
        //public Transportadoras Transporadora { get; set; }
        public string Transporadora { get; set; }
        public string NumeroGuia { get; set; }
        public string CodigoConvenio { get; set; }
        public string DepartamentoRemitente { get; set; }
        public string CiudadRemitente { get; set; }
        public string NombreRemitente { get; set; }
        public string EmailRemitente { get; set; }
        public string DireccionRemitente { get; set; }
        public string TelefonoRemitente { get; set; }

        public string DepartamentoDestino { get; set; }
        public string CiudadDestino { get; set; }
        public string NombreDestino { get; set; }
        public string EmailDestino { get; set; }
        public string DireccionDestino { get; set; }
        public string BarrioDestino { get; set; }
        public string TelefonoDestino { get; set; }

        public string Observaciones { get; set; }
        public string ValorPagar { get; set; }

    }
}

