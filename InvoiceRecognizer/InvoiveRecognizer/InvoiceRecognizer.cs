using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.CognitiveServices.FormRecognizer;
using Microsoft.Azure.CognitiveServices.FormRecognizer.Models;
using Newtonsoft.Json;

namespace InvoiceRecognizer
{
    public class InvoiceRecognizer : CodeActivity
    {
        [Category("Input")]
        public InArgument<string> SubscriptionKey { get; set; }

        [Category("Input")]
        public InArgument<string> FormRecognizerEndPoint { get; set; }

        [Category("Input")]
        public InArgument<string> FormRecognizerModelId { get; set; }

        [Category("Input")]
        public InArgument<string> FilePath { get; set; }

        [Category("Output")]
        public OutArgument<string> InvoiceJson { get; set; }

        private readonly Dictionary<string, string> FormOutputToInvoice =  new Dictionary<string, string>();

        private void SetupDictionaryMapping()
        {
            FormOutputToInvoice.Add("DATE:", "InvoiceDate");
            FormOutputToInvoice.Add("INVOICE#:", "InvoiceNumber");
            FormOutputToInvoice.Add("FROM:", "From");
            FormOutputToInvoice.Add("TO:", "To");
            FormOutputToInvoice.Add("TERMS:", "Terms");
            FormOutputToInvoice.Add("DUE:", "Due");
            FormOutputToInvoice.Add("Subtotal", "Subtotal");
            FormOutputToInvoice.Add("Tax", "TaxPercent");
            FormOutputToInvoice.Add("BALANCE DUE", "BalanceDue");
            FormOutputToInvoice.Add("S.No.", "SerialNumber");
            FormOutputToInvoice.Add("Item Description", "ItemDescription");
            FormOutputToInvoice.Add("Quantity", "Quantity");
            FormOutputToInvoice.Add("Price", "Price");
            FormOutputToInvoice.Add("Amount", "Amount");
        }

        protected override void Execute(CodeActivityContext context)
        {

            // From Local - manual entries
            //  var jsonResult = GetInvoiceDetailsfromSample();


            // From Azure Form Recognizer
            SetupDictionaryMapping();
            string subscriptionKey = SubscriptionKey.Get(context);
            string formRecognizerEndPoint = FormRecognizerEndPoint.Get(context);
            string formRecognizerModelId = FormRecognizerModelId.Get(context);
            string filePath = FilePath.Get(context);

            string jsonResult = GetInvoiceDetailsfromAzureFormRecognizer(subscriptionKey, formRecognizerEndPoint, 
                                                                        formRecognizerModelId, filePath);
            // Set the output
            InvoiceJson.Set(context, jsonResult);
        }       

        private string GetInvoiceDetailsfromAzureFormRecognizer(string subscriptionKey, string formRecognizerEndPoint,
                                                                        string formRecognizerModelId, string filePath)
        {
            // Create form client object with Form Recognizer subscription key
            IFormRecognizerClient formClient = new FormRecognizerClient(
                new ApiKeyServiceClientCredentials(subscriptionKey))
            {
                Endpoint = formRecognizerEndPoint
            };

            Console.WriteLine("Analyze PDF form...");
            string jsonResult =  AnalyzePdfForm(formClient, new Guid(formRecognizerModelId), filePath);

            return jsonResult;
           

        }

        private string AnalyzePdfForm(IFormRecognizerClient formClient, Guid formRecognizerModelId, string filePath)
        {

            //TODO: As this is part of .NET Library, writing to console does not makes sense
            // Write to an Exceptions section either in jsonResult or another output and send it to caller
            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("\nInvalid file path.");
                return null;
            }

           string jsonResult = "";
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open))
                {
                    //Get the result from an asynchronous call but as a synchnous way by using .Result
                    AnalyzeResult result = formClient.AnalyzeWithCustomModelAsync(formRecognizerModelId, stream, contentType: "application/pdf").Result;
                    jsonResult = TransformAnalyzeResult(result);
                }
            }
            catch (ErrorResponseException e)
            {
                Console.WriteLine("Analyze PDF form : " + e.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Analyze PDF form : " + ex.Message);
            }

            return jsonResult;

        }

        private string TransformAnalyzeResult(AnalyzeResult result)
    {
        Invoice inv = new Invoice();

        foreach (var page in result.Pages)
        {
                // Take the page from the FormRecognizer output and convert that into InvoicePage object
                inv.invoicePages.Add(GetInvoicePage(page));
        }

        string jsonResult = JsonConvert.SerializeObject(inv);
        return jsonResult;
    }

        private InvoicePage GetInvoicePage(ExtractedPage page)
    {
        InvoicePage invPage = new InvoicePage()
        {
            PageNumber = page.Number
        };
        // Take the KeyValuePairs from page (returned by Form Recognizer) and set the properties of InvoicePage object
        GetInvoiceKeyValuePairs(page, invPage);
        // Take the Tables from page (returned by Form Recognizer) and assign each table info to InvoiceLineItem object 
        // and add that line item object to Invoice Page line item collection
        GetInvoiceTables(page, invPage);

        return invPage;
    }

        private void GetInvoiceKeyValuePairs(ExtractedPage page, InvoicePage invPage)
        {
            foreach (var kv in page.KeyValuePairs)
            {
                if (kv.Key.Count > 0)
                {
                    if (FormOutputToInvoice.ContainsKey(kv.Key[0].Text))
                    {

                        string dictionaryValue = FormOutputToInvoice[kv.Key[0].Text];

                        // Normally we will have to write a humoungous switch statement to set the value for the correspoding property
                        // Instead, we are takning advantage of reflection to set the value to the property
                         /*   switch (dictionaryValue)
                            {
                                case "InvoiceDate":
                                    invPage.InvoiceDate = kv.Value[0].Text;
                                    break;
                                case "InvoiceNumber":
                                    invPage.InvoiceNumber = kv.Value[0].Text;
                                    break;
                                // and SO ON
                                default:
                                    break;
                            }

                        */

                        // Use Reflection to get a reference to the property based on name of the property as a string
                        PropertyInfo piInstance = typeof(InvoicePage).GetProperty(FormOutputToInvoice[kv.Key[0].Text]);
                        if (kv.Value.Count > 0)
                        {
                            //TODO: Right now, I am changing all properties to string type to test.
                            // Change these data types back and see how to type cast kv.Value[0].Text to the correct datatype and set the value

                            // Concatenate the values if more than one value is there for a key, for example FROM and TO
                            // using a linq expression here to concatenate the values instead of a for loop

                           /* StringBuilder sb = new StringBuilder();
                            foreach (var s in kv.Value )
                            {
                                sb.Append(s.Text);
                            }
                            piInstance.SetValue(invPage, sb.ToString(), null);
                            */

                            piInstance.SetValue(invPage, string.Join(", ", kv.Value.Select(v => v.Text)), null);
                            
                            // If there is only one value for a given key, the following code would have been sufficient
                            //piInstance.SetValue(invPage, kv.Value[0].Text, null);
                        }
                    }
                }
            }
        }

        private void GetInvoiceTables(ExtractedPage page, InvoicePage invPage)
        {
            foreach (var t in page.Tables)
            {
                // Each column will have the same number of entries. Take the entries in the first column (any column is OK)
                int numberOfRows = ((t.Columns.Count > 0 && t.Columns[0].Entries.Count > 0) ? t.Columns[0].Entries.Count : 0);
            
                if (numberOfRows > 0)
                {
                    //Create number of InvoiceLineItems corresponding to number of rows in the table
                      /*  List<InvoiceLineItem> invoiceLineItems = new List<InvoiceLineItem>();
                        for (int i = 0; i < numberOfRows; i++)
                        {
                            invoiceLineItems.Add(new InvoiceLineItem());
                        }
                        */

                    List<InvoiceLineItem> invoiceLineItems = Enumerable.Range(0, numberOfRows)
                                                .Select(i => new InvoiceLineItem())
                                                .ToList();

                    // Fill each row
                    foreach (var c in t.Columns)
                    {
                        if (FormOutputToInvoice.ContainsKey(c.Header[0].Text))
                        {
                            PropertyInfo piInstance = typeof(InvoiceLineItem).GetProperty(FormOutputToInvoice[c.Header[0].Text]);
                            for (int i = 0; i < numberOfRows; i++)
                            {
                                // As each entry contains only one value, we take the first value
                                piInstance.SetValue(invoiceLineItems[i], c.Entries[i][0].Text, null);

                            }
                        }
                    }
                    invPage.invoiceLineItems = invoiceLineItems;
                }
            }
        }

        private object GetInvoiceDetailsfromSample()
        {
            Invoice inv = new Invoice();
            InvoicePage invPage = new InvoicePage()
            {
                InvoiceDate = "2018/12/24",
                InvoiceNumber = "544",
                Terms = "NET 30",
                Due = "2019/01/24",
                Subtotal = "3000.00",
                TaxPercent = "6.50",
                TaxAmount = "195.00",
                BalanceDue = "3195.00"
            };

       
            //First line
            InvoiceLineItem invLineItem = new InvoiceLineItem() {
                SerialNumber = "1",
                ItemDescription = "Item 1",
                Quantity = "2",
                Price = "600.00",
                Amount = "1200.00"

                /*SerialNumber = 1,
                ItemDescription = "Item 1",
                Quantity = 2,
                Price = 600.00,
                Amount = 1200.00
                */
            };
            invPage.invoiceLineItems.Add(invLineItem);


            // Second Line
            InvoiceLineItem invLineItem2 = new InvoiceLineItem() {
                SerialNumber = "2",
                ItemDescription = "Item 2",
                Quantity = "4",
                Price = "450.00",
                Amount = "1800.00"

                /*SerialNumber = 2,
                ItemDescription = "Item 2",
                Quantity = 4,
                Price = 450.00,
                Amount = 1800.00
                */
            };
            invPage.invoiceLineItems.Add(invLineItem2);


            inv.invoicePages.Add(invPage);

            var jsonResult = JsonConvert.SerializeObject(inv);

            return jsonResult;

        }
    }

    public class Invoice
    {
        public IList<InvoicePage> invoicePages = new List<InvoicePage>();
    }

    public class InvoicePage
    {

        //TODO: Uncomment later to the correct data types 
        // when the code in GetInvoiceKeyValuePairs is fixed.
        public string InvoiceDate { get; set; }
        public string InvoiceNumber { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Terms { get; set; }
        public string Due { get; set; }
        public string Subtotal { get; set; }
        public string TaxPercent { get; set; }
        public string TaxAmount { get; set; }
        public string BalanceDue { get; set; }
        public int? PageNumber { get; set; }

       /* public DateTime Invoicedate { get; set; }
        public int InvoiceNumber { get; set; }
        public string Terms { get; set; }
        public DateTime Due { get; set; }
        public double Subtotal { get; set; }
        public double TaxPercent { get; set; }
        public double TaxAmount { get; set; }
        public double BalanceDue { get; set; }
        public int? PageNumber { get; set; }
        */

        public IList<InvoiceLineItem> invoiceLineItems = new List<InvoiceLineItem>();
    }

    public class InvoiceLineItem
    {
        //TODO: Uncomment later to the correct data types 
        // when the code in GetInvoiceTables is fixed.
        public string SerialNumber { get; set; }
        public string ItemDescription { get; set; }
        public string Quantity { get; set; }
        public string Price { get; set; }
        public string Amount { get; set; }


        /*public int SerialNumber { get; set; }
        public string ItemDescription { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
        public double Amount { get; set; }
        */
    }

    /* Useful links
     * https://docs.microsoft.com/en-us/dotnet/standard/design-guidelines/capitalization-conventions
     * https://docs.microsoft.com/en-us/dotnet/api/system.reflection.propertyinfo.setvalue?view=netframework-4.8
     * 
     */
}

