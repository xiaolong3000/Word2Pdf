using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using System.IO;
using Aspose.Words.Saving;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string WordPath;
        string OutPath;
        string WordsPath;
        string OutsPath;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public static void WtoP(string w,string o)//string wordInputPath, string pdfOutputPath)
        {
            
            try {
                Aspose.Words.Document doc = new Aspose.Words.Document(w);
                if (doc == null) { throw new Exception("Word文件无效或者Word文件被加密！"); }
                doc.Save(o);




            }catch(Exception e){
                MessageBox.Show("首先选择Word文件位置，再选择转换成PDF文件位置，输入PDF名字，最后再点击转换", "出现错误", MessageBoxButtons.OK, MessageBoxIcon.Information);
           // throw e;
                ActiveForm.Close();
              
                 
           
            }  



        }
        public static void WstoPs(string w, string o)//string wordInputPath, string pdfOutputPath)
        {

        
              try
                {
                    if (!System.IO.Directory.Exists(w))
                        throw new System.IO.DirectoryNotFoundException();
            string[] fileNames = null;
           

            fileNames = System.IO.Directory.GetFiles(w, "*.doc", System.IO.SearchOption.AllDirectories);
            foreach (string name in fileNames)
            {
              //  Console.WriteLine(name);
                string[] namebysplite = Regex.Split(name, "\\\\", RegexOptions.IgnoreCase);
              

                    Aspose.Words.Document doc = new Aspose.Words.Document(name);
                    if (doc == null) { throw new Exception("Word文件无效或者Word文件被加密！"); }
                    string s=namebysplite[namebysplite.Length-1];
                    doc.Save(Path.Combine(o,s.Substring(0,s.Length-4)+".pdf"),SaveFormat.Pdf);


                
            }


                }
              catch (Exception e)
              {
                  MessageBox.Show("首先选择Word文件位置，再选择转换成PDF文件位置，输入PDF名字，最后再点击转换", "出现错误", MessageBoxButtons.OK, MessageBoxIcon.Information);
                  // throw e;
                  ActiveForm.Close();
              }
               


/*

            try
            {
              
                Aspose.Words.Document doc = new Aspose.Words.Document(w);
                if (doc == null) { throw new Exception("Word文件无效或者Word文件被加密！"); }
                doc.Save(o);


            }
            catch (Exception e)
            {
                throw e;
            }

*/

        }

      


        private void button1_Click(object sender, EventArgs e)
        {
            WtoP(WordPath,OutPath);
            label1.Text = "Word转换为PDF完成";
            label1.Visible = true;
        }

        private void 关闭ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog word = new OpenFileDialog();
            word.Filter = "Word文件|*.doc";

            if (word.ShowDialog() == DialogResult.OK)
            { WordPath = word.FileName;
            label2.Text = WordPath;
            label2.Visible = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog outPath=new SaveFileDialog();
            outPath.Filter = "PDF文件|*.pdf";
            if (outPath.ShowDialog() == DialogResult.OK)
            { OutPath = outPath.FileName;
            label3.Text = OutPath;
            label3.Visible = true;
            
            
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog word = new FolderBrowserDialog();
            if (word.ShowDialog() == DialogResult.OK) {
                WordsPath = word.SelectedPath.ToString();
                label4.Text = WordsPath;
                label4.Visible = true;
                
            
            }
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog word = new FolderBrowserDialog();
            if (word.ShowDialog() == DialogResult.OK)
            {
                OutsPath = word.SelectedPath.ToString();
                label5.Text = OutsPath;
                label5.Visible = true;

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            WstoPs(WordsPath, OutsPath);
            label6.Text = "批量转换完成";
            label6.Visible = true;
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

     

      

      

      
    

       
    }

    


}
