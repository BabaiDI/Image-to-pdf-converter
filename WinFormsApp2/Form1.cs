using Aspose.Words;

namespace WinFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new()
            {
                Multiselect = true
            };
            if (openFile.ShowDialog() != DialogResult.OK)
                return;

            Document doc = new();
            DocumentBuilder builder = new(doc);
            progressBar1.Maximum = openFile.FileNames.Length;
            try
            {
                foreach (string Name in openFile.FileNames)
                {
                    builder.InsertImage(Name);
                    progressBar1.Value += 1;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return;
            }

            SaveFileDialog save = new()
            {
                Filter = "* | *.pdf"
            };
            if (save.ShowDialog() == DialogResult.OK)
            {
                doc.Save(save.FileName);
                MessageBox.Show("Saved");
            }
            progressBar1.Value = 0;
        }
    }
}