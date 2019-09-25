using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InvoiceRecognizer;

namespace TestCognitiveInvoice
{
    class Program
    {
        static void Main(string[] args)
        {
            Activity invRecognizer = new InvoiceRecognizer.InvoiceRecognizer();

            //TODO: Bring these values from a local config file
            IDictionary<string, object> input = new Dictionary<string, object> {
                                                        { "SubscriptionKey", "<SUBSCRIPTION KEY>" },
                                                        { "FormRecognizerEndPoint", "<END POINT URL>" },
                                                        { "FormRecognizerModelId", "<MODEL ID>" },
                                                        { "FilePath", @"<FILE PATH>" }
                                                };
          
            IDictionary<string, object> output = WorkflowInvoker.Invoke(invRecognizer, input);
            Console.WriteLine("Analyze PDF form...");
        }
    }
}
