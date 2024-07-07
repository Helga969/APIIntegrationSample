using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dadata;
using Dadata.Model;
using System.Windows.Forms;
using Word=Microsoft.Office.Interop.Word;

namespace iis_lr1
{

    public partial class Form1 : Form
    {
        public Form1()
        {   
            InitializeComponent();
        }
        Word._Application Word = new Word.Application();
        
        private Word._Document GetDoc(string path)
        {
            Word._Document doc = Word.Documents.Add(path);
            SetTemplate(doc);
            return doc;
        }


        private void SetTemplate(Word._Document fs)
        {
            fs.Bookmarks["organization1"].Range.Text = label7.Text;
            fs.Bookmarks["dayMountYear"].Range.Text = DateTime.Today.ToShortDateString();
            fs.Bookmarks["organization2"].Range.Text = label8.Text;
            fs.Bookmarks["FIO1"].Range.Text = FIO_1.Text;
            fs.Bookmarks["FIO2"].Range.Text = FIO_2.Text;
            fs.Bookmarks["doc1"].Range.Text = Doc_1.Text;
            fs.Bookmarks["doc2"].Range.Text = Doc_2.Text;
            fs.Bookmarks["city"].Range.Text = City.Text;
            fs.Bookmarks["inn_1"].Range.Text = Inn_1.Text;
            fs.Bookmarks["inn_2"].Range.Text = Inn_2.Text;
            fs.Bookmarks["cpp_1"].Range.Text = Cpp_1.Text;
            fs.Bookmarks["cpp_2"].Range.Text = Cpp_2.Text;
            fs.Bookmarks["post_index_pokup"].Range.Text = Post_in_1.Text;
            fs.Bookmarks["post_index_prod"].Range.Text = Post_in_2.Text;
            fs.Bookmarks["pokup_adres"].Range.Text = Adress_1.Text;
            fs.Bookmarks["prod_adres"].Range.Text = Adress_2.Text;
            if (radioButton1.Checked == true)
            {
                fs.Bookmarks["tovar"].Range.Text = Tovar.Text;
                fs.Bookmarks["adres"].Range.Text = Adress_deliv.Text;
            }
            if (radioButton2.Checked == true)
            {
                fs.Bookmarks["predoplataDay"].Range.Text = timePred.Text;
                fs.Bookmarks["predoplata"].Range.Text = procPred.Text;
                fs.Bookmarks["raschet"].Range.Text = raschet.Text;
            }

                
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (FIO_1.Text==string.Empty || FIO_2.Text == string.Empty || Doc_1.Text == string.Empty || Doc_2.Text == string.Empty || City.Text == string.Empty || Inn_1.Text == string.Empty || Inn_2.Text == string.Empty || Cpp_1.Text == string.Empty || Cpp_2.Text == string.Empty || Post_in_1.Text == string.Empty || Post_in_2.Text == string.Empty || Adress_1.Text == string.Empty || Adress_2.Text == string.Empty)
            {
                MessageBox.Show("Поля пустые", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                if (radioButton1.Checked == true)
                {
                    
                    if (Tovar.Text == string.Empty || Adress_deliv.Text == string.Empty)
                    {
                        MessageBox.Show("Поля пустые", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        Word._Document doc = GetDoc(Environment.CurrentDirectory + "\\postavka.docx");
                        Object newDoc = Path.Combine(doc.Path, DateTime.Now.ToString() + doc.Name);
                        doc.Close();
                    }
                }
                if (radioButton2.Checked == true)
                {
                    if(timePred.Text == string.Empty || procPred.Text == string.Empty || raschet.Text == string.Empty)
                    {
                        MessageBox.Show("Поля пустые", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else 
                    {
                        Word._Document doc = GetDoc(Environment.CurrentDirectory + "\\predoplata.docx");
                        Object newDoc = Path.Combine(doc.Path, DateTime.Now.ToString() + doc.Name);
                        doc.Close(); 
                    }
                    
                }
            }
            
            
       }
        
        private void button2_Click(object sender, EventArgs e)
        {
            var token = ConfigurationManager.AppSettings["Token"];
            var api = new SuggestClient(token);//SuggestClientAsync(token, secret);
            var result = api.SuggestParty(Name_org_1.Text);
            string outName = String.Format("{0}", result.suggestions[0].value);
            label7.Text = outName;
            string outFIO = String.Format("{0}", result.suggestions[0].data.management.name);
            FIO_1.Text = outFIO;
            string outUrAdr = String.Format("{0}", result.suggestions[0].data.address.value);
            Adress_1.Text = outUrAdr;
            
            string outINN = String.Format("{0}", result.suggestions[0].data.inn);
            Inn_1.Text = outINN;
            string outKPP = String.Format("{0}", result.suggestions[0].data.kpp);
            Cpp_1.Text = outKPP;
            string outInd = String.Format("{0}", result.suggestions[0].data.address.data.postal_code);
            Post_in_1.Text = outInd;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            var token = ConfigurationManager.AppSettings["Token"];
            var api = new SuggestClient(token);//SuggestClientAsync(token, secret);
            var result = api.SuggestParty(Name_org_2.Text);
            string outName = String.Format("{0}", result.suggestions[0].value);
            label8.Text = outName;
            string outFIO = String.Format("{0}", result.suggestions[0].data.management.name);
            FIO_2.Text = outFIO;
            string outUrAdr = String.Format("{0}", result.suggestions[0].data.address.value);
            Adress_2.Text = outUrAdr;
            string outINN = String.Format("{0}", result.suggestions[0].data.inn);
            Inn_2.Text = outINN;
            string outKPP = String.Format("{0}", result.suggestions[0].data.kpp);
            Cpp_2.Text = outKPP;
            string outInd = String.Format("{0}", result.suggestions[0].data.address.data.postal_code);
            Post_in_2.Text = outInd;
            string outUrCity = String.Format("{0}", result.suggestions[0].data.address.data.city);
            City.Text = outUrCity;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            var token = ConfigurationManager.AppSettings["Token"];
            var secret = ConfigurationManager.AppSettings["Secret"];
            var api = new CleanClient(token, secret);
            var result = api.Clean<Address>(add_proverka.Text);
            string add_deliv = String.Format("{0}", result.result);
            Adress_deliv.Text = add_deliv;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                Tovar.Enabled=true;
                add_proverka.Enabled = true;
                Adress_deliv.Enabled = true;
                button4.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                Tovar.Enabled = false;
                add_proverka.Enabled = false;
                Adress_deliv.Enabled = false;
                button4.Enabled = false;
                button1.Enabled = false;
            }    
            
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                timePred.Enabled = true;
                procPred.Enabled = true;
                raschet.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                timePred.Enabled = false;
                procPred.Enabled = false;
                raschet.Enabled = false;
                button1.Enabled = false;
            }

        }
      
    }
}

