using System;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace Docx_to_pdf_convertor
{
    public partial class Form1 : Form
    {
        public Document wordDocument { get; set; }
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory="C:\\";
            openFileDialog1.Filter = "Docx files (*.doc)|*.docx|All files (*.*)|*.*";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string docfilename= openFileDialog1.FileName;
                    textBox1.Text = docfilename;
                    Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                    wordDocument = appWord.Documents.Open(docfilename);
                    wordDocument.Activate();
                    wordDocument.SaveAs2(@"D:\myfile.pdf",WdSaveFormat.wdFormatPDF);
                    wordDocument.Close();

                }
                catch (Exception ex)
                { }
            }

        }
    }
}
