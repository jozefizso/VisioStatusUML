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
        private IVisio.Shape DesenharForma(IVisio.Page page, StatusSinistro sinistro)
        {
            var shape = page.DrawRectangle(1, 1, 3, 2);
            shape.Text = string.IsNullOrEmpty(sinistro.NomeStatusSeguinte) ? "Não achei" : sinistro.NomeStatusSeguinte;
            shape.Data1 = sinistro.CodigoStatusAtual;
            shape.Data2 = sinistro.CodigoStatusSeguinte;
            shape.Data3 = sinistro.CodigoTransacao;

            return shape;
        }
        private void ConnectarDesenhos(IVisio.Shape shapeAtual, IVisio.Shape shapeAnterior)
        {
            shapeAtual.AutoConnect(shapeAnterior, IVisio.VisAutoConnectDir.visAutoConnectDirUp);
        }

        private void DesenharPainel(IVisio.Page page, StatusSinistro sinistro, List<StatusSinistro> listaSinistro, List<IVisio.Shape> listaShape)
        {
            if (!listaShape.Any(x => x.Data2 == sinistro.CodigoStatusSeguinte))
            {
                var novoShape = DesenharForma(page, sinistro);
                listaShape.Add(novoShape);

                foreach(var s in listaShape.Where(x => x.Data2 == novoShape.Data1))
                {
                    ConnectarDesenhos(novoShape, s);
                }

                foreach(var item in listaSinistro.Where(x => !listaShape.Any(y => y.Data3 == x.CodigoTransacao) && x.CodigoStatusAtual == sinistro.CodigoStatusSeguinte))
                {
                    DesenharPainel(page, item, listaSinistro, listaShape);
                }
            }
            else
            {
                var shapeAtual = listaShape.Single(x => x.Data2 == sinistro.CodigoStatusAtual);
                var shapeSeguinte = listaShape.Single(x => x.Data2 == sinistro.CodigoStatusSeguinte);
                ConnectarDesenhos(shapeAtual, shapeSeguinte);
            }
        }
        private void GerarArquivo(List<StatusSinistro> lista)
        {
            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");
            var page = visapp.ActivePage;

            List<IVisio.Shape> listaShape = new List<IVisio.Shape>();

            DesenharPainel(page, lista.FirstOrDefault(), lista, listaShape);
        }
        private List<StatusSinistro> MontarListaDeStatus(List<string[]> lista)
        {
            List<StatusSinistro> s = new List<StatusSinistro>();

            lista.ForEach(x => s.Add(new StatusSinistro()
            {
                CodigoTransacao = x[0],
                NomeTransacao = x[1],

                CodigoStatusAtual = x[2],
                NomeStatusAtual = x[3],

                CodigoStatusSeguinte = x[4],
                NomeStatusSeguinte = x[5]
            }));

            StatusSinistro primeiroStatus = s.SingleOrDefault(x => x.CodigoTransacao == "A" || x.CodigoTransacao == "00A");
            List<StatusSinistro> statusPosteriores = s.Where(x => x.CodigoStatusAtual != null && x.CodigoStatusSeguinte != null).ToList();
            List<StatusSinistro> outrosStatus = s.Where(x => (x.CodigoStatusAtual == null || x.CodigoStatusSeguinte == null) && x.CodigoTransacao != "A").ToList();

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
