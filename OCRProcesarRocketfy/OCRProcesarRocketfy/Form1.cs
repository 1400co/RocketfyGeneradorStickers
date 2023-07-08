

using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace OCRProcesarRocketfy
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            var pedidos = new List<Pedidos>();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                pedidos = LeerPedidosDeExcel(filePath);
            }

            if (pedidos.Count == 0)
            {
                MessageBox.Show("No hay pedidos en estado pendiente.");
                return;
            }

            dgPedidos.DataSource = pedidos;

            var exeLocation = AppDomain.CurrentDomain.BaseDirectory;
            var pdfFilePath = System.IO.Path.Combine(exeLocation, "pedidos.pdf");

            this.CreatePdfReport(pedidos, pdfFilePath);

            // Para abrir el archivo después de generarlo
            System.Diagnostics.Process.Start(pdfFilePath);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            var pedidos = new List<Pedidos>();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                pedidos = LeerPedidosDeExcel(filePath);
            }

            if (pedidos.Count == 0)
            {
                MessageBox.Show("No hay pedidos en estado pendiente.");
                return;
            }

            dgPedidos.DataSource = pedidos;

            var exeLocation = AppDomain.CurrentDomain.BaseDirectory;
            var pdfFilePath = System.IO.Path.Combine(exeLocation, "pedidos.pdf");

            this.CreatePdfReportCarta(pedidos, pdfFilePath);

            // Para abrir el archivo después de generarlo
            System.Diagnostics.Process.Start(pdfFilePath);
        }

        public List<Pedidos> LeerPedidosDeExcel(string filePath)
        {
            var pedidos = new List<Pedidos>();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // Comenzar en la segunda fila (ignorar la primera que es de encabezados)
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 2].Value.ToString() == "No ha sido generada")
                        continue;

                    if (worksheet.Cells[row, 10].Value.ToString() != "Pendiente" && worksheet.Cells[row, 10].Value.ToString() != "En empaque")
                        continue;

                    var pedido = new Pedidos
                    {
                        Transporadora = worksheet.Cells[row, 30].Value.ToString(), // Asumiendo que la columna 'Transportadora' es un enum
                        NumeroGuia = worksheet.Cells[row, 2].Value.ToString(),
                        CodigoConvenio = "901195703-4 (87622)", // Este valor no se encuentra en el excel según lo proporcionado
                        DepartamentoRemitente = "SANTANDER", // Este valor no se encuentra en el excel según lo proporcionado
                        CiudadRemitente = "FLORIDABLANCA", // Este valor no se encuentra en el excel según lo proporcionado
                        NombreRemitente = $"Natutrend ({worksheet.Cells[row, 31].Value})",
                        EmailRemitente = worksheet.Cells[row, 13].Value.ToString(),
                        DireccionRemitente = "Cra 6 # 7 - 06 apto 403 edificio rayenaris", // Este valor no se encuentra en el excel según lo proporcionado
                        TelefonoRemitente = "3188426287",
                        DepartamentoDestino = worksheet.Cells[row, 17].Value.ToString(),
                        CiudadDestino = worksheet.Cells[row, 16].Value.ToString(),
                        NombreDestino = worksheet.Cells[row, 11].Value.ToString(),
                        EmailDestino = worksheet.Cells[row, 13].Value.ToString(),
                        DireccionDestino = worksheet.Cells[row, 15].Value.ToString(),
                        BarrioDestino = worksheet.Cells[row, 18].Value.ToString(),
                        TelefonoDestino = worksheet.Cells[row, 12].Value.ToString(),
                        Observaciones = worksheet.Cells[row, 20].Value.ToString(), // Este valor no se encuentra en el excel según lo proporcionado
                        ValorPagar = "$" + worksheet.Cells[row, 25].Value.ToString(),
                    };

                    pedidos.Add(pedido);
                }
            }

            return pedidos;
        }


        private void ExtractTextFromPdf(string filePath)
        {
            //try
            //{

            //    List<Pedidos> listaPedidos = new List<Pedidos>();

            //    using (PdfReader reader = new PdfReader(filePath))
            //    {
            //        int pageCount = reader.NumberOfPages;

            //        for (int i = 1; i <= pageCount; i++)
            //        {
            //            string pageText = PdfTextExtractor.GetTextFromPage(reader, i);
            //            var result = ProcessPageText(pageText);
            //            if (result != null) { listaPedidos.Add(result); }
            //        }
            //    }

            //    dgPedidos.DataSource = listaPedidos;

            //    MessageBox.Show("Extracción de texto y creación de reporte completadas.");

            //    MessageBox.Show("Extracción de texto completada.");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error al extraer texto del PDF: " + ex.Message);
            //}
        }

        public void CreatePdfReport(List<Pedidos> listaPedidos, string pdfPath)
        {
            // Calcula el tamaño de la página en puntos (paisaje)
            Rectangle pageSize = new Rectangle(ConvertToPoint(10), ConvertToPoint(8));

            // Calcula los márgenes en puntos
            float margin = ConvertToPoint(0.5f);  // 5 mm
            float marginL = ConvertToPoint(0.2f);  // 5 mm

            using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
            {
                Document document = new Document(pageSize, ConvertToPoint(0.2f), ConvertToPoint(1.5f), ConvertToPoint(1f), ConvertToPoint(0.2f));
                PdfWriter writer = PdfWriter.GetInstance(document, fs);

                document.Open();

                foreach (var pedido in listaPedidos)
                {
                    PdfPTable table = new PdfPTable(2);
                    table.WidthPercentage = 90;
                    float[] columnWidths = { 1f, 1f };
                    table.SetWidths(columnWidths);

                    PdfPCell cell;

                    cell = new PdfPCell(new Phrase($"Transportadora: {pedido.Transporadora}\nNúmero de guía: {pedido.NumeroGuia}\nCódigo de convenio: {pedido.CodigoConvenio}\n", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    cell.Colspan = 2;
                    table.AddCell(cell);

                    // Añade el código de barras
                    Barcode128 barcode = new Barcode128
                    {
                        Code = pedido.NumeroGuia,
                        BarHeight = 20f,  // ajusta la altura del código de barras
                        X = 1.5f,  // ajusta el ancho del código de barras
                    };
                    iTextSharp.text.Image barcodeImage = barcode.CreateImageWithBarcode(writer.DirectContent, null, null);
                    PdfPCell barcodeCell = new PdfPCell(barcodeImage, fit: false);
                    barcodeCell.Colspan = 2;
                    barcodeCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    barcodeCell.PaddingBottom = 1f;  // agrega un espacio debajo del código de barras
                    table.AddCell(barcodeCell);

                    cell = new PdfPCell(new Phrase($"Remitente: {pedido.NombreRemitente} ({pedido.EmailRemitente})\nDirección remitente: {pedido.DireccionRemitente}, {pedido.CiudadRemitente}, {pedido.DepartamentoRemitente}\nTeléfono remitente: {pedido.TelefonoRemitente}", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase($"Destinatario: {pedido.NombreDestino} ({pedido.EmailDestino})\nDirección de destino: {pedido.DireccionDestino}, {pedido.CiudadDestino}, {pedido.DepartamentoDestino}\nTeléfono de destino: {pedido.TelefonoDestino}", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase($"Observaciones: {pedido.Observaciones}\nValor a pagar: {pedido.ValorPagar}", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    cell.Colspan = 2;
                    table.AddCell(cell);

                    document.Add(table);
                    document.NewPage(); // crea una nueva página para cada guía de transporte
                }

                document.Close();
            }
        }

        public void CreatePdfReportCarta(List<Pedidos> listaPedidos, string pdfPath)
        {
            // Define el tamaño de la página (Carta)
            Rectangle pageSize = PageSize.Letter;

            // Calcula los márgenes en puntos
            float margin = ConvertToPoint(0.5f);  // 5 mm
            float marginL = ConvertToPoint(0.2f);  // 5 mm

            using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
            {
                Document document = new Document(pageSize, ConvertToPoint(0.2f), ConvertToPoint(1.5f), ConvertToPoint(1f), ConvertToPoint(0.2f));
                PdfWriter writer = PdfWriter.GetInstance(document, fs);

                document.Open();

                // Crear el objeto MultiColumnText
                MultiColumnText multiColumnText = new MultiColumnText();
                multiColumnText.AddRegularColumns(document.Left, document.Right, 10f, 2);  // 10f es el espacio entre las columnas y 2 es el número de columnas

                foreach (var pedido in listaPedidos)
                {
                    PdfPTable table = new PdfPTable(2);
                    table.WidthPercentage = 100;  // La tabla cubrirá el ancho completo de la columna
                    float[] columnWidths = { 1f, 1f };
                    table.SetWidths(columnWidths);

                    PdfPCell cell;

                    cell = new PdfPCell(new Phrase($"Transportadora: {pedido.Transporadora}\nNúmero de guía: {pedido.NumeroGuia}\nCódigo de convenio: {pedido.CodigoConvenio}\n", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    cell.Colspan = 2;
                    table.AddCell(cell);

                    // Añade el código de barras
                    Barcode128 barcode = new Barcode128
                    {
                        Code = pedido.NumeroGuia,
                        BarHeight = 20f,  // ajusta la altura del código de barras
                        X = 1.5f,  // ajusta el ancho del código de barras
                    };
                    iTextSharp.text.Image barcodeImage = barcode.CreateImageWithBarcode(writer.DirectContent, null, null);
                    PdfPCell barcodeCell = new PdfPCell(barcodeImage, fit: false);
                    barcodeCell.Colspan = 2;
                    barcodeCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    barcodeCell.PaddingBottom = 1f;  // agrega un espacio debajo del código de barras
                    table.AddCell(barcodeCell);

                    cell = new PdfPCell(new Phrase($"Remitente: {pedido.NombreRemitente} ({pedido.EmailRemitente})\nDirección remitente: {pedido.DireccionRemitente}, {pedido.CiudadRemitente}, {pedido.DepartamentoRemitente}\nTeléfono remitente: {pedido.TelefonoRemitente}", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase($"Destinatario: {pedido.NombreDestino} ({pedido.EmailDestino})\nDirección de destino: {pedido.DireccionDestino}, {pedido.CiudadDestino}, {pedido.DepartamentoDestino}\nTeléfono de destino: {pedido.TelefonoDestino}", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase($"Observaciones: {pedido.Observaciones}\nValor a pagar: {pedido.ValorPagar}", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8)));
                    cell.Colspan = 2;
                    table.AddCell(cell);

                    table.KeepTogether = true;  // Asegura que la tabla (pedido) no se divide entre páginas
                    multiColumnText.AddElement(table);

                    // Agrega un espacio después de cada pedido
                    Paragraph paragraph = new Paragraph();
                    AddEmptyLine(paragraph, 1);
                    multiColumnText.AddElement(paragraph);
                }

                document.Add(multiColumnText);
                document.Close();
            }
        }

        // Función auxiliar para agregar líneas vacías
        private static void AddEmptyLine(Paragraph paragraph, int number)
        {
            for (int i = 0; i < number; i++)
            {
                paragraph.Add(new Paragraph(" "));
            }
        }

        // Función auxiliar para convertir cm a puntos
        private float ConvertToPoint(float cm)
        {
            return cm * 28.35f;
        }

        private Pedidos ProcessPageText(string pageText)
        {
            var result = this.AnalizarPagina(pageText);

            ProcesarGuias procesarGuias = new ProcesarGuias();

            switch (result)
            {
                case Transportadoras.Tcc:
                    return procesarGuias.ProcesarTCC(pageText);
                    break;
                case Transportadoras.Servientrega:
                    return procesarGuias.ProcesarServiEntrega(pageText);
                    break;
                case Transportadoras.Interrapidisimo:
                    return procesarGuias.ProcesarInter(pageText);
                    break;
                case Transportadoras.Ninguna:
                    return null;
                    break;
                default:
                    return null;
                    break;
            }


        }

        private Transportadoras AnalizarPagina(string textoOcr)
        {
            if (textoOcr.Contains("GLOBAL MENSAJERIA S.A.S"))
                return Transportadoras.Tcc;

            if (textoOcr.Contains("Servientrega S.A NIT"))
                return Transportadoras.Servientrega;

            if (textoOcr.Contains("www.interrapidisimo.com"))
                return Transportadoras.Interrapidisimo;


            return Transportadoras.Ninguna;
        }


    }
}
