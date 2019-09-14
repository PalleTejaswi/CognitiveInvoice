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
            invPage.fromCompanyName = "Apple";
            invPage.fromCompanyAddress = "1 Infinite Loop, Cupertino, CA 95014";

            //First line
            InvoiceLineItem invLineItem = new InvoiceLineItem();
            invLineItem.serialNumber = 1;
            invLineItem.itemDescription = "Item 1";
            invPage.invoiceLineItems.Add(invLineItem);

            // Second Line
            InvoiceLineItem invLineItem2 = new InvoiceLineItem();
            invLineItem2.serialNumber = 2;
            invLineItem2.itemDescription = "Item 2";
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
        public string fromCompanyName { get; set; }
        public string fromCompanyAddress { get; set; }
        

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

