using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OCRProcesarRocketfy
{
    public class GeneradorFormatos
    {

        public void GenerarFormatoDespacho(List<Pedidos> listaPedidos, string pdfPath)
        {
            // Agrupar los pedidos por transportadora
            var pedidosPorTransportadora = listaPedidos.GroupBy(p => p.Transporadora)
                                                      .OrderBy(g => g.Key)
                                                      .ToList();

            // Crear el documento PDF
            Document document = new Document(PageSize.Letter);

            // Crear el escritor de PDF
            using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
            {
                PdfWriter writer = PdfWriter.GetInstance(document, fs);

                // Abrir el documento
                document.Open();

                foreach (var grupoPedidos in pedidosPorTransportadora)
                {
                    // Agregar el título de la transportadora
                    Paragraph tituloTransportadora = new Paragraph($"Relación de despacho - Transportadora: {grupoPedidos.Key} - Total de pedidos: {grupoPedidos.Count()}");
                    tituloTransportadora.Alignment = Element.ALIGN_CENTER;
                    tituloTransportadora.SpacingAfter = 10;
                    tituloTransportadora.Font.Size = 10;
                    document.Add(tituloTransportadora);

                    // Agregar la fecha actual
                    Paragraph fechaActual = new Paragraph($"Fecha: {DateTime.Now.ToShortDateString()}");
                    fechaActual.Alignment = Element.ALIGN_RIGHT;
                    fechaActual.SpacingAfter = 10;
                    fechaActual.Font.Size = 8;
                    document.Add(fechaActual);

                    // Crear la tabla de pedidos
                    PdfPTable table = new PdfPTable(7);
                    table.WidthPercentage = 100;
                    float[] columnWidths = { 1f, 1f, 1f, 1f, 1f, 1f, 1f };
                    table.SetWidths(columnWidths);

                    // Agregar encabezados de columna
                    PdfPCell headerCell = new PdfPCell(new Phrase("Código Rocketfy", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    headerCell.BackgroundColor = BaseColor.LightGray;
                    table.AddCell(headerCell);

                    headerCell = new PdfPCell(new Phrase("Número de Guía", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    headerCell.BackgroundColor = BaseColor.LightGray;
                    table.AddCell(headerCell);

                    headerCell = new PdfPCell(new Phrase("Código de Convenio", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    headerCell.BackgroundColor = BaseColor.LightGray;
                    table.AddCell(headerCell);

                    headerCell = new PdfPCell(new Phrase("Nombre del Destinatario", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    headerCell.BackgroundColor = BaseColor.LightGray;
                    table.AddCell(headerCell);

                    headerCell = new PdfPCell(new Phrase("Ciudad del Destinatario", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    headerCell.BackgroundColor = BaseColor.LightGray;
                    table.AddCell(headerCell);

                    headerCell = new PdfPCell(new Phrase("Departamento del Destinatario", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    headerCell.BackgroundColor = BaseColor.LightGray;
                    table.AddCell(headerCell);

                    headerCell = new PdfPCell(new Phrase("Nombre del Producto", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    headerCell.BackgroundColor = BaseColor.LightGray;
                    table.AddCell(headerCell);

                    // Agregar filas de pedidos
                    foreach (var pedido in grupoPedidos)
                    {
                        table.AddCell(new Phrase(pedido.CodigoRocket, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                        table.AddCell(new Phrase(pedido.NumeroGuia, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                        table.AddCell(new Phrase(pedido.CodigoConvenio, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                        table.AddCell(new Phrase(pedido.NombreDestino, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                        table.AddCell(new Phrase(pedido.CiudadDestino, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                        table.AddCell(new Phrase(pedido.DepartamentoDestino, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                        table.AddCell(new Phrase(pedido.Observaciones, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    }

                    // Agregar la tabla al documento
                    document.Add(table);

                    // Agregar el espacio para la firma
                    Paragraph espacioFirma = new Paragraph("\n\n\n\n\n\n\n\nRecibido: ___________________________");
                    espacioFirma.Alignment = Element.ALIGN_LEFT;
                    espacioFirma.Font.Size = 8;
                    document.Add(espacioFirma);

                    // Agregar una nueva página
                    document.NewPage();
                }

                // Cerrar el documento
                document.Close();
            }
        }
    }
}
