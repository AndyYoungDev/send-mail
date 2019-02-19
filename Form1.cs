using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 群发
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;

            InitializeComponent();

        }
        private void sendMail()
        {
            string localPath = System.IO.Directory.GetCurrentDirectory();
            StreamReader sr = new StreamReader(localPath+"/发件html", System.Text.Encoding.UTF8);
            string restOfStream = sr.ReadToEnd();
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                try
                {
                    MailMessage message = new MailMessage();
                    MailAddress fromAddr = new MailAddress("邮箱");
                    message.From = fromAddr;
                    message.To.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());
                    message.Subject = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    message.IsBodyHtml = true;

                    message.Attachments.Add(new Attachment(localPath + "/a.png"));
                    message.Attachments[0].ContentType.Name = "image/gif";
                    message.Attachments[0].ContentId = "a";
                    message.Attachments[0].ContentDisposition.Inline = true;
                    message.Attachments[0].TransferEncoding = System.Net.Mime.TransferEncoding.Base64;

                    //<img src="cid:a">



                    message.Body = restOfStream;


                    message.Attachments.Add(new Attachment("文件"));
                    SmtpClient client = new SmtpClient("smtp地址", 25);
                    client.EnableSsl = false;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new NetworkCredential("邮箱", "密码");

                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.Send(message);
                    dataGridView1.Rows[i].Cells[3].Value = "已发送";
                    dataGridView1.Rows[i].Cells[4].Value = DateTime.Now.ToString(); ;
                }
                catch (Exception ee)
                {
                    dataGridView1.Rows[i].Cells[3].Value = "发送失败";
                }
            }
            button2.Enabled = true;
            button1.Enabled = true;
            button2.Text = "开始发送";
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog pOpenFileDialog = new OpenFileDialog();
            pOpenFileDialog.Filter = "所有文件|*.txt";
            pOpenFileDialog.Multiselect = false;
            pOpenFileDialog.Title = "打开文件";
            if (pOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = pOpenFileDialog.FileName;
                this.readFile(path);
            }
        }
        private void readFile(string filename)
        {
            StreamReader sr = new StreamReader(filename, Encoding.Default);
            String line;
            ArrayList data = new ArrayList();
            while ((line = sr.ReadLine()) != null)
            {
                data.Add(line);
            }
            for (int i = 0; i < data.Count; i++)
            {
                string[] sArray = Regex.Split(data[i].ToString(), "----", RegexOptions.IgnoreCase);

                if (sArray[0] !=""&& sArray[1]!="") {
                    int index = this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[index].Cells[0].Value = i + 1;
                    this.dataGridView1.Rows[index].Cells[1].Value = sArray[0];
                    this.dataGridView1.Rows[index].Cells[2].Value = sArray[1];
                    this.dataGridView1.Rows[index].Cells[3].Value = "未发送";
                    this.dataGridView1.Rows[index].Cells[4].Value = "无";
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Text = "正在发送";
            button2.Enabled = false;
            button1.Enabled = false;
            Task.Factory.StartNew(sendMail);
        }
    }
}
