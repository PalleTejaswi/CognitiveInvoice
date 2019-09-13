using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace InvoiceRecognizer
{
    public class InvoiceRecognizer : CodeActivity
    {
        [Category("Input")]
        public InArgument<string> filePath { get; set; }

        [Category("Output")]
        public OutArgument<string> invoiceJson { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            // Call 

            /* if (kv.Key.Count > 0)
                 if (kv.Key[0].Text == "Address")
                     invPage.fromCompanyAddress = kv.Value[0].Text;
            */

            Invoice inv = new Invoice();
            InvoicePage invPage = new InvoicePage();
            invPage.fromCompanyName = "Apple";
            invPage.fromCompanyAddress = "1 Infinite Loop, Cupertino, CA 95014";

            //First line
            InvoiceLineItem invLineItem = new InvoiceLineItem();
            invLineItem.invoiceNumber = 123456;
            invLineItem.invoiceDate = new DateTime(2019, 9, 12);
            invPage.invoiceLineItems.Add(invLineItem);

            // Second Line
            InvoiceLineItem invLineItem2 = new InvoiceLineItem();
            invLineItem2.invoiceNumber = 765433;
            invLineItem2.invoiceDate = new DateTime(2019, 9, 11);
            invPage.invoiceLineItems.Add(invLineItem2);


            inv.invoicePages.Add(invPage);

            var jsonResult = JsonConvert.SerializeObject(inv);

            invoiceJson.Set(context, jsonResult);

            //put a for loop here
        }
    }

    public class Invoice
    {
        public IList<InvoicePage> invoicePages = new List<InvoicePage>();

    }

    public class InvoicePage
    {
        public string fromCompanyName { get; set; }
        public string fromCompanyAddress { get; set; }
        // Add more keys here

        public IList<InvoiceLineItem> invoiceLineItems = new List<InvoiceLineItem>();

    }

    public class InvoiceLineItem
    {
        public long invoiceNumber { get; set; }
        public DateTime invoiceDate { get; set; }

        // Add more columns here
    }

}

