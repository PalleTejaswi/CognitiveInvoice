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

            
            //TODO: Test with manual entries initially. Once UiPath is working properly, change this code 
            // to get the invoice data from Azure Form Recognizer

            Invoice inv = new Invoice();
            InvoicePage invPage = new InvoicePage();
            //TODO: Change the following code to reflect the new properties
            invPage.date = new DateTime(2018, 12, 24); ;
            invPage.invoiceNumber = 544;
            invPage.terms = "NET 30";
            invPage.due = new DateTime(2019, 01, 24);
            invPage.subtotal = 3000.00;
            invPage.taxPercent = 6.50 ;
            invPage.taxAmount = 195.00;
            invPage.balanceDue = 3195.00;

            //First line
            InvoiceLineItem invLineItem = new InvoiceLineItem();
            invLineItem.serialNumber = 1;
            invLineItem.itemDescription = "Item 1";
            invLineItem.quantity = 2;
            invLineItem.price = 600.00;
            invLineItem.amount = 1200.00;
            invPage.invoiceLineItems.Add(invLineItem);

            // Second Line
            InvoiceLineItem invLineItem2 = new InvoiceLineItem();
            invLineItem2.serialNumber = 2;
            invLineItem2.itemDescription = "Item 2";
            invLineItem.quantity = 4;
            invLineItem.price = 450.00;
            invLineItem.amount = 1800.00;
            invPage.invoiceLineItems.Add(invLineItem2);


            inv.invoicePages.Add(invPage);

            var jsonResult = JsonConvert.SerializeObject(inv);

            invoiceJson.Set(context, jsonResult);
        }
    }

    public class Invoice
    {
        public IList<InvoicePage> invoicePages = new List<InvoicePage>();

    }

    public class InvoicePage
    {

        //TODO: Change these keys based on new invoice
        public DateTime date { get; set; }
        public int invoiceNumber { get; set; }

        public string terms { get; set; }
        public DateTime due { get; set; }
        public double subtotal { get; set; }
        public double taxPercent { get; set; }

        public double taxAmount { get; set; }
        public double balanceDue { get; set; }

        public IList<InvoiceLineItem> invoiceLineItems = new List<InvoiceLineItem>();

    }

    public class InvoiceLineItem
    {
        public int serialNumber { get; set; }
        public string itemDescription { get; set; }
        public int quantity { get; set; }
        public double price { get; set; }
        public double amount { get; set; }
    }
}

