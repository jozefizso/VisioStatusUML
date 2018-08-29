using System;
using System.IO;
using System.Windows.Forms;
using VisioAutomation.Extensions;
using Excel = Microsoft.Office.Interop.Excel;
using VisioAutomation.Geometry;

using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Msagl.Layout.LargeGraphLayout;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Models.Dom;
using VisioStatusUML.Dominio;

namespace VisioStatusUML
{
    public partial class frmXMLVisio : Form
    {
        private readonly OpenFileDialog dialog;
        private Document doc;

        public frmXMLVisio()
        {
            InitializeComponent();
            dialog = new OpenFileDialog();
        }

        private void btnVisio_Click(object sender, EventArgs e)
        {
            var lista = LerArquivo();
            var sinistro = MontarListaDeStatus(lista);

            GerarArquivo(sinistro);
        }



        private void DesenharStatus(IVisio.Page page, StatusSinistro sinistro, List<StatusSinistro> listaSinistro, List<IVisio.Shape> listaShape)
        {
            //todo: Ajustar metodo de desenho para criar o shape de todos os status relacionados a um item antes de ir para o nivel abaixo. Para evitar erro de reutilização da mesma transação.


            if (listaShape != null)
            {
                if (listaShape.Any(x => x.Data2 == sinistro.CodigoStatusSeguinte))
                {
                    var shape = listaShape.SingleOrDefault(x => x.Data2 == sinistro.CodigoStatusSeguinte);

                    foreach (var s in listaShape.Where(x => x.Data2 == sinistro.CodigoStatusAnterior))
                    {
                        s.AutoConnect(shape, IVisio.VisAutoConnectDir.visAutoConnectDirDown);
                    }

                    listaShape.Add(shape);
                }
                else
                {
                    var shape = page.DrawRectangle(1, 1, 3, 2);
                    shape.Text = string.IsNullOrEmpty(sinistro.NomeStatusSeguinte) ? "Não achei" : sinistro.NomeStatusSeguinte;
                    shape.Data1 = sinistro.CodigoStatusAnterior;
                    shape.Data2 = sinistro.CodigoStatusSeguinte;
                    shape.Data3 = sinistro.CodigoTransacao;

                    if (listaShape.Any(x => x.Data2 == sinistro.CodigoStatusAnterior))
                    {
                        foreach (var s in listaShape.Where(x => x.Data2 == sinistro.CodigoStatusAnterior))
                        {
                            s.AutoConnect(shape, IVisio.VisAutoConnectDir.visAutoConnectDirDown);
                        }
                    }

                    listaShape.Add(shape);
                }

                var novaLista = listaSinistro.Where(x => x.CodigoTransacao != sinistro.CodigoTransacao).ToList();

                if (novaLista.Any(x => x.CodigoStatusAnterior == sinistro.CodigoStatusSeguinte))
                {
                    foreach (var item in novaLista.Where(x => x.CodigoStatusAnterior == sinistro.CodigoStatusSeguinte))
                    {
                        DesenharStatus(page, item, novaLista, listaShape);
                    }
                }
            }
        }
        private void GerarArquivo(List<StatusSinistro> lista)
        {
            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");
            var page = visapp.ActivePage;

            List<IVisio.Shape> listaShape = new List<IVisio.Shape>();

            DesenharStatus(page, lista.FirstOrDefault(), lista, listaShape);
        }
        private List<StatusSinistro> MontarListaDeStatus(List<string[]> lista)
        {
            List<StatusSinistro> s = new List<StatusSinistro>();

            lista.ForEach(x => s.Add(new StatusSinistro()
            {
                CodigoTransacao = x[0],
                NomeTransacao = x[1],

                CodigoStatusAnterior = x[2],
                NomeStatusAnterior = x[3],

                CodigoStatusSeguinte = x[4],
                NomeStatusSeguinte = x[5]
            }));

            StatusSinistro primeiroStatus = s.SingleOrDefault(x => x.CodigoTransacao == "A" || x.CodigoTransacao == "00A");
            List<StatusSinistro> statusPosteriores = s.Where(x => x.CodigoStatusAnterior != null && x.CodigoStatusSeguinte != null).ToList();
            List<StatusSinistro> outrosStatus = s.Where(x => (x.CodigoStatusAnterior == null || x.CodigoStatusSeguinte == null) && x.CodigoTransacao != "A").ToList();

            List<StatusSinistro> sinistros = new List<StatusSinistro> { primeiroStatus };
            sinistros.AddRange(statusPosteriores);

            return sinistros;
        }
        private List<string[]> LerArquivo()
        {
            List<string[]> contents = new List<string[]>();
            dialog.Filter = "Excel|*.xlsx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Excel.Application xl = new Excel.Application();
                Workbook workbook = xl.Workbooks.Open(dialog.FileName);
                Worksheet sheet = workbook.Sheets[1];

                int numRows = sheet.UsedRange.Rows.Count;
                int numColumns = 7;     // according to your sample

                for (int rowIndex = 2; rowIndex <= numRows; rowIndex++)  // assuming the data starts at 1,1
                {
                    string[] record = new string[numColumns];

                    for (int colIndex = 1; colIndex <= numColumns; colIndex++)
                    {
                        Range cell = (Range)sheet.Cells[rowIndex, colIndex];
                        if (cell.Value != null && Convert.ToString(cell.Value) != null)
                        {
                            string c = Convert.ToString(cell.Value);
                            record[colIndex - 1] = c == "NULL" ? null : c;
                        }
                    }

                    if (record.Where(x => x == null).Count() != 7)
                        contents.Add(record);
                    else
                        break;
                }

                xl.Quit();
                Marshal.ReleaseComObject(xl);
            }

            return contents;
        }
    }
}
