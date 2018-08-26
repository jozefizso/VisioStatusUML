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

            List<StatusSinistro> statusIniciais = s.Where(x => x.CodigoStatusAnterior == null).ToList();
            List<StatusSinistro> statusIntermediarios = s.Where(x => x.CodigoStatusAnterior != null && x.CodigoStatusSeguinte != null).ToList();
            List<StatusSinistro> statusTerminais = s.Where(x => x.CodigoStatusSeguinte == null).ToList();

            List<StatusSinistro> sinistros = new List<StatusSinistro>();

            sinistros.AddRange(statusIniciais);

            foreach(var item in sinistros)
            {
                item.ListaStatusSeguintes = new List<StatusSinistro>();
                item.ListaStatusSeguintes.AddRange(statusIntermediarios.Where(x => x.CodigoStatusAnterior == item.CodigoStatusSeguinte));

                foreach (var item2 in item.ListaStatusSeguintes)
                {
                    item2.ListaStatusSeguintes = new List<StatusSinistro>();
                    item2.ListaStatusSeguintes.AddRange(statusTerminais.Where(x => x.CodigoStatusAnterior == item2.CodigoStatusSeguinte));
                }
            }

            //sinistros = statusIniciais.FirstOrDefault();
            //sinistros.ListaStatusSeguintes = new List<StatusSinistro>();
            //sinistros.ListaStatusSeguintes.Add(statusIntermediarios.FirstOrDefault());
            //sinistros.ListaStatusSeguintes.FirstOrDefault().ListaStatusSeguintes = new List<StatusSinistro>();
            //sinistros.ListaStatusSeguintes.FirstOrDefault().ListaStatusSeguintes.Add(statusTerminais.FirstOrDefault());


            return sinistros;
        }
        private void GerarArquivo(List<StatusSinistro> listaNivel1)
        {
            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");
            var page = visapp.ActivePage;

            foreach (var item in listaNivel1)
            {
                DesenharStatus(page, item);
            }

        }

        private void DesenharStatus(IVisio.Page page, StatusSinistro sinistro, IVisio.Shape shapeAnterior = null)
        {
            var shape = page.DrawRectangle(1, 1, 3, 2);
            shape.Text = string.IsNullOrEmpty(sinistro.NomeStatusSeguinte) ? sinistro.NomeStatusAnterior : sinistro.NomeStatusSeguinte;

            if (shapeAnterior != null)
                shapeAnterior.AutoConnect(shape, IVisio.VisAutoConnectDir.visAutoConnectDirDown);

            if(sinistro.ListaStatusSeguintes != null && sinistro.ListaStatusSeguintes.Count > 0)
            {
                foreach(var item in sinistro.ListaStatusSeguintes)
                {
                    DesenharStatus(page, item, shape);
                }
            }
        }
    }
}
