using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Services.Description;
using System.Windows.Forms;

namespace OCRProcesarRocketfy
{
    public class ProcesarGuias
    {
        public ProcesarGuias()
        {

        }

        public Pedidos ProcesarTCC(string textoOcr)
        {
            var result = new Pedidos() { Transporadora = "TCC" };

            try
            {
                string[] lineas = textoOcr.Split('\n');


                result.NumeroGuia = this.ObtenerValorRemesaTCC(textoOcr);
                result.CodigoConvenio = this.PartirCadena(lineas[5], '-', 0);

                result.NombreRemitente = this.PartirCadena(lineas[5], '-', 1);
                result.DepartamentoRemitente = this.ObtenerDepartamentoTCC(lineas[7]);
                result.CiudadRemitente = this.ObtenerMunicipioTCC(lineas[7]);
                result.TelefonoRemitente = lineas.Count() == 19 ? this.ObtenerPrimerNumeroTCC(lineas[13]) : this.ObtenerPrimerNumeroTCC(lineas[12]);
                result.DireccionRemitente = lineas[9];

                result.NombreDestino = this.PartirCadena(lineas[6], '-', 1);
                result.TelefonoDestino = this.PartirCadena(lineas[6], '-', 0);


                if (this.ObtenerDepartamentoTCC(lineas[8]) == "")
                    result.DepartamentoDestino = this.ObtenerDepartamentoTCC(lineas[10]);
                else
                    result.DepartamentoDestino = this.ObtenerDepartamentoTCC(lineas[8]);

                if (this.ObtenerMunicipioTCC(lineas[8]) == "")
                    result.CiudadDestino = this.ObtenerMunicipioTCC(lineas[10]);
                else
                    result.CiudadDestino = this.ObtenerMunicipioTCC(lineas[8]);

                result.BarrioDestino = lineas.Count() == 19 ? lineas[12] : "";
                result.DireccionDestino = lineas.Count() == 19 ? lineas[9] : lineas[10];

                result.Observaciones = lineas.Count() == 19 ? lineas[14] : lineas[13];



            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return result;

        }

        public Pedidos ProcesarServiEntrega(string textoOcr)
        {
            var result = new Pedidos() { Transporadora = "Servientrega" };

            try
            {
                string[] lineas = textoOcr.Split('\n');


                result.NumeroGuia = this.ObtenerValorPatron(textoOcr, @"GUIA No.\s+(\d+)");
                result.CodigoConvenio = this.ObtenerValorPatron(textoOcr, @"CÓDIGO SER: (.+)");
                result.NombreRemitente = EncontrarFilaPatron(lineas, "GUIA No.");
                result.DepartamentoRemitente = "SANTANDER";
                result.CiudadRemitente = "FLORIDABLANCA";
                result.TelefonoRemitente = lineas[29];
                result.DireccionRemitente = "CRA 6 # 7 - 06 APTO 403";

                result.NombreDestino = this.EncontrarFilaPatron(lineas, "Nombre");
                result.TelefonoDestino = this.PartirCadena(this.EncontrarFilaPatron(lineas, 50, 10, "No reclamado Teléfono:"), ':', 1);
                result.DepartamentoDestino = this.EncontrarFilaPatron(lineas, 28, 10, "CREDITO").Replace("CREDITO", "");
                result.CiudadDestino = this.ObtenerValorPatron(textoOcr, @"CIUDAD: (.+)");
                result.BarrioDestino = "";
                result.DireccionDestino = this.EncontrarFilaPatron(lineas, 100, 15, "/ / /");

                result.Observaciones = this.EncontrarFilaPatron(lineas, 60, 15, "Obs. para Entrega:"); // lineas.Count() != 219 ? this.PartirCadena(lineas[64], ':', 1) : this.PartirCadena(lineas[63], ':', 1);
                result.ValorPagar = this.EncontrarFilaPatron(lineas, "Vr. a Cobrar:"); //this.EncontrarFilaPatron(lineas, 70, 15, "Vr. Total:");//lineas.Count() == 222 ? lineas[80] : lineas[79]; "Vr. Total:


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return result;
        }

        public Pedidos ProcesarInter(string textoOcr)
        {
            var result = new Pedidos() { Transporadora = "Interrapidisimo" };

            try
            {
                string[] lineas = textoOcr.Split('\n');


                result.NumeroGuia = this.ObtenerValorPatron(textoOcr, @"No.\s+(\d+)");
                result.CodigoConvenio = this.ObtenerValorPatron(lineas[18], @"\b\d+\b") + " - " + "Nit 901195703";
                var nombres = this.EncontrarFilaPatron(lineas, "DESTINATARIO REMITENTE");
                result.NombreRemitente = this.PartirCadena(nombres, '-', 1); ;

                result.DepartamentoRemitente = this.PartirCadena(lineas[22], '\\', 1);
                result.CiudadRemitente = this.PartirCadena(lineas[22], '\\', 0);
                result.TelefonoRemitente = this.PartirCadena(lineas[19], ' ', 1);
                result.DireccionRemitente = "cra 6 # 7 - 06 Ed Rayenaris apto 403";

                result.NombreDestino = this.PartirCadena(nombres, '-', 0);
                result.TelefonoDestino = this.PartirCadena(lineas[19], ' ', 0); ;
                result.DepartamentoDestino = this.PartirCadena(lineas[4], '\\', 1);
                result.CiudadDestino = this.PartirCadena(lineas[4], '\\', 0);
                result.BarrioDestino = "";

                var direccion = this.EncontrarFilaPatron(lineas, "NIT");

                result.DireccionDestino = this.PartirCadena(direccion, ',', 0);

                result.Observaciones = lineas[33];
                result.ValorPagar = lineas[31];


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return result;
        }

        private string ObtenerPrimerNumeroTCC(string cadena)
        {
            string patron = @"\d+";
            Match match = Regex.Match(cadena, patron);

            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return string.Empty;
            }
        }

        private string ObtenerValorRemesaTCC(string texto)
        {
            string valorRemesa = string.Empty;

            // Utilizar expresión regular para buscar el valor después de "REMESA:"
            string patron = @"REMESA:\s+(\d+)";
            Match match = Regex.Match(texto, patron);

            if (match.Success)
            {
                valorRemesa = match.Groups[1].Value;
            }

            return valorRemesa;
        }

        private string ObtenerDepartamentoTCC(string textoOcr)
        {
            string patron = @"\((.*?)\)";
            Match match = Regex.Match(textoOcr, patron);

            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            else
            {
                return string.Empty;
            }
        }

        private string ObtenerMunicipioTCC(string cadena)
        {
            string patron = @"^(.*?)\s+\(";
            Match match = Regex.Match(cadena, patron);

            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            else
            {
                return string.Empty;
            }
        }

        private string ObtenerValorPatron(string texto, string patron)
        {
            string valorRemesa = string.Empty;

            // Utilizar expresión regular para buscar el valor después de "REMESA:"
            Match match = Regex.Match(texto, patron);

            if (match.Success)
            {
                valorRemesa = match.Groups[1].Value == "" ? match.Value : match.Groups[1].Value;
            }

            return valorRemesa;
        }

        private string EncontrarFilaPatron(string[] lista, int fila, int intentos, string patron)
        {
            if (intentos == 0)
                return "";

            while (intentos > 0)
            {
                if (lista[fila].Contains(patron))
                    return lista[fila];

                fila++;
                intentos--;
            }

            return "";
        }

        private string EncontrarFilaPatron(string[] lista, string llave)
        {
            int posicion = Array.FindIndex(lista, item => item.Contains(llave));

            return lista[posicion + 1];
        }

        private string PartirCadena(string cadena, char caracter, int posicion)
        {
            if (!cadena.Contains(caracter))
                return cadena;

            string[] lineas = cadena.Split(caracter);
            return lineas[posicion];
        }
    }
}
