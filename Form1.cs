﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;

namespace mail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            ImapClient client = new ImapClient();
            client.Host = "imap.yandex.ru";  //"imap.mail.ru"
            client.Port = 993;
            client.Username = "testLibrary123@yandex.ru"; //"librarytest123@mail.ru"
            client.Password = "TestTest123"; // "support321"
            client.SecurityOptions = SecurityOptions.SSLImplicit;

            ImapQueryBuilder imapQueryBuilder = new ImapQueryBuilder(Encoding.Default);
            
            imapQueryBuilder.Subject.Contains("поддержка"); //necessarily writing in Russian, as it does not work without encoding (Unsupporting encoding)
            
            imapQueryBuilder.InternalDate.On(Convert.eTime("31.03.2020"));
                

            MailQuery query;
            client.ReadOnly = true; 
            client.SelectFolder("Inbox");

            ImapMessageInfoCollection messages = new ImapMessageInfoCollection();

            try
            {
                query = imapQueryBuilder.GetQuery();
                messages = client.ListMessages(query);
                MessageBox.Show("Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unsuccessfully: " + ex);
            }
            button1.Enabled = true;
        }
        
    }
}
