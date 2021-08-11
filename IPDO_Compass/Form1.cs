using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using IPDO_Compass.DB;
using System.Data.Common;
using System.Threading;
using Ookii.Dialogs;

namespace IPDO_Compass
{
    public partial class Form1 : Form
    {
        int id_ipdo = 0;
        double[,] termo = new double[5, 13];

        public Form1()
        {

            InitializeComponent();


            /* if (File.Exists(@"P:\RISCO\IPDO\11_2019\IPDO-21-11-2019.xlsm"))
             {
                 string PathFileName = @"P:\RISCO\IPDO\11_2019\IPDO-21-11-2019.xlsm";

                 Program.carregaIPDO(PathFileName);
                 Program.GerTerm(PathFileName);
                 Program.IPDO_Sub(PathFileName);

             }*/



        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public void SelecionarSaida()
        {

            //FolderBrowserDialog d = new FolderBrowserDialog();
            var data = DateTime.Today;
            var d = new OpenFileDialog();
           
            d.Filter = "Excel (*.xlsm)|*.xlsm|All files (*.*)|*.*";

            var IPDO = d.FileName;
            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txt_Caminho.Text = d.FileName;
                if (txt_Caminho.Text != "")
                {
                    try
                    {
                        var ipdoArq = txt_Caminho.Text;
                        //Program.carregaIPDO(ipdoArq);
                        //Program.GerTerm(ipdoArq);
                        //Program.IPDO_Sub(ipdoArq);

                        Program.carregaIPDO(ipdoArq, "azure");
                        Program.GerTerm(ipdoArq, "azure");
                        Program.IPDO_Sub(ipdoArq, "azure");
                        MessageBox.Show("IPDO carregado com sucesso!");

                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString());
                    }
                }

            }
            else
            {
                return;

            }

            //Ookii.Dialogs.WinForms.VistaFolderBrowserDialog ofd = new Ookii.Dialogs.WinForms.VistaFolderBrowserDialog();
            //ofd.SelectedPath = @"P:\RISCO\IPDO\";

            //if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    txt_Caminho.Text = ofd.SelectedPath;
            //}



        }

        private void bt_seleciona_Click(object sender, EventArgs e)
        {
            SelecionarSaida();

        }
    }
}
