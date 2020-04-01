using System;
using System.Collections.Generic;
using Microsoft.ML;
using System.Linq;
using Microsoft.ML.Runtime.Data;
using Microsoft.ML.Runtime.Api;
using System.IO;
using EAGetMail;
using System.Globalization;

namespace ML_Based_Invoice_Prediction
{

    /// <summary>
    /// User defined Datatype used for Email Input and Output
    /// </summary>
    public class Model_Inputs_Outputs
    {
        //Email Subject text as string
        [Column(ordinal: "0")]
        public string EmailSubject { get; set; }

        //Boolean variable to check if Email is an invoice or not
        [Column(ordinal: "1", name: "Label")]
        public bool IsInvoice { get; set; }
    }

    /// <summary>
    /// User defined Datatype used for Predicting if email subject is invoice of not
    /// </summary>
    public class Model_Predictions
    {
        //Boolean variable to check if Email is an invoice or not
        [Column("2", "PredictedLabel")]
        public bool IsInvoice { get; set; }
    }

    /// <summary>
    /// Class used to perform all Model training and predictions
    /// </summary>
    public class ML_Model
    {
        //Save email in Guest Directory

        static string _generateFileName(int sequence)
        {
            DateTime currentDateTime = DateTime.Now;
            return string.Format("{0}-{1:000}-{2:000}.eml",
                currentDateTime.ToString("yyyyMMddHHmmss", new CultureInfo("en-US")),
                currentDateTime.Millisecond,
                sequence);
        }
        //List of training data points
        static List<Model_Inputs_Outputs> trainingData = new List<Model_Inputs_Outputs>();

        //List of test data points
        static List<Model_Inputs_Outputs> testData = new List<Model_Inputs_Outputs>();

        /// <summary>
        /// Loads the training data using the sample Email subjects
        /// </summary>
        static void LoadTrainingData()
        {
            //Reading the email sample subjects from txt file to list of strings
            List<string> lines = new List<string>();
            lines = File.ReadAllLines(GlobalVariables.trainingDataFilePath).ToList();

            //Loop to split subject and invoice boolean using | and add to training data set
            foreach (string line in lines)
            {
                bool isItInvoice = false;
                string[] partsOfLine = line.Split('|');
                if (partsOfLine[1].ToLower() == "yes")
                    isItInvoice = true;

                trainingData.Add(new Model_Inputs_Outputs()
                {
                    EmailSubject = partsOfLine[0],
                    IsInvoice = isItInvoice
                });

            }
        }

        /// <summary>
        /// Loads the test data using more sample email subjects
        /// </summary>
        static void LoadTestData()
        {
            //Reading the email sample subjects from txt file to list of strings
            List<string> lines = new List<string>();
            lines = File.ReadAllLines(GlobalVariables.testDataFilePath).ToList();

            //Loop to split subject and invoice boolean using | and add to test data set
            foreach (string line in lines)
            {
                bool isItInvoice = false;
                string[] partsOfLine = line.Split('|');
                if (partsOfLine[1].ToLower() == "yes")
                    isItInvoice = true;

                testData.Add(new Model_Inputs_Outputs()
                {
                    EmailSubject = partsOfLine[0],
                    IsInvoice = isItInvoice
                });

            }
        }

        /// <summary>
        /// Method used for training and executing the ML Model to get prediction of inovice
        /// </summary>
        public void Execute_ML_Model()
        {
            int i = 0;
            //Loop until user chooses to exit the program
            while (true)
            {
                //Call Methods to load training and test data 
                LoadTrainingData();
                LoadTestData();


                //Use ML Context to implement Model 
                MLContext mlContext = new MLContext();

                //Transform data from list to IDataView
                IDataView trainingDataView = mlContext.CreateStreamingDataView<Model_Inputs_Outputs>(trainingData);

                //Define a pipeline so that model uses featurized input
                var pipeline = mlContext.Transforms.Text.FeaturizeText("EmailSubject", "Features").
                    Append(mlContext.BinaryClassification.Trainers.FastTree(numLeaves: 50, numTrees: 50, minDatapointsInLeaves: 1));

                //Input training data into the model
                var model = pipeline.Fit(trainingDataView);

                //Testing model accuracy by validating using test data set
                IDataView testDataView = mlContext.CreateStreamingDataView<Model_Inputs_Outputs>(testData);
                var predictions = model.Transform(testDataView);
                var metrics = mlContext.BinaryClassification.Evaluate(predictions, "Label");
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.Title = "Machine Learning POC";
                System.Globalization.NumberFormatInfo nfi = new System.Globalization.CultureInfo("en-US", false).NumberFormat;
                nfi.PercentDecimalDigits = 0;
                Console.WriteLine("Accuracy of the Model = " + metrics.Accuracy.ToString("P", nfi));



                //User inputs email subject or Exit to stop run
                //Commenting the below lines of code to retrieve email from Outlook 365 - Subrato - Start
                //Console.WriteLine("Enter an Email Subject or Enter Exit to Terminate Program : ");
                //Commenting the below lines of code to retrieve email from Outlook 365 - Subrato - End
                Console.WriteLine("Enter any Key to Continue or Enter Exit to Terminate Program : ");
                string userInputStringForExit = Console.ReadLine();

                //Break loop and stop run if Exit is used
                if (userInputStringForExit.ToLower() == "exit")
                break;


                //Use model to make prediction

                //Outlook.Application app = new Outlook.Application();
                //Outlook.NameSpace outlookNs = app.GetNamespace("MAPI");
                //Outlook.MAPIFolder emailFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

                //List<MailItem> ReceivedEmail = new List<MailItem>();
                //foreach (Outlook.MailItem mail in emailFolder.Items)
                //    ReceivedEmail.Add(mail);

                //foreach (MailItem mail in ReceivedEmail)
                //{
                //    //do stuff
                //}


                MailServer oServer = new MailServer("imap-mail.outlook.com",
                    "subrato.biswas@exlservice.com", "Password", ServerProtocol.Imap4);

                oServer.SSLConnection = true;

                // Set 993 SSL port
                oServer.Port = 993;

                MailClient oClient = new MailClient("TryIt");
                oClient.Connect(oServer);

                MailInfo[] infos = oClient.GetMailInfos();
                Console.WriteLine("Total {0} email(s)\r\n", infos.Length);

                
                    MailInfo info = infos[i];
                    Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);

                    // Receive email from IMAP4 server
                    Mail oMail = oClient.GetMail(info);

                    if (oMail.From.ToString().Equals(""))
                    {

                    }
                    string userInputString = oMail.Subject;

                    // Generate an unqiue email file name based on date time.
                    

                    // Save email to local disk
                    //oMail.SaveAs(fullPath, true);

                    // Mark email as deleted from IMAP4 server.
                    //oClient.Delete(info);
                

                // Quit and expunge emails marked as deleted from IMAP4 server.
                
            

            var predictionFunction = model.MakePredictionFunction
                                              <Model_Inputs_Outputs, Model_Predictions>(mlContext);

                Model_Inputs_Outputs inputToModel = new Model_Inputs_Outputs();

                inputToModel.EmailSubject = userInputString;

                var invoicePrediction = predictionFunction.Predict(inputToModel);

                //Display on terminal if email is invoice or not
                if (invoicePrediction.IsInvoice)
                    Console.WriteLine("This is an Invoice");

                else
                    Console.WriteLine("This is NOT an Invoice");

                //Add the user input subject to the training data set
                Console.WriteLine("Was the prediction correct Y/N?");

                string userResponse = Console.ReadLine();
                Console.WriteLine(Environment.NewLine);

                string addNewEmailSubject;

                //Logic to add correct invoice classification for email subject
                if (userResponse.ToLower() == "y" && invoicePrediction.IsInvoice)
                    addNewEmailSubject = userInputString + "|Yes";
                else if (userResponse.ToLower() == "y" && !invoicePrediction.IsInvoice)
                    addNewEmailSubject = userInputString + "|No";
                else if (userResponse.ToLower() != "y" && invoicePrediction.IsInvoice)
                    addNewEmailSubject = userInputString + "|No";
                else
                    addNewEmailSubject = userInputString + "|Yes";

                //Adding subject to training data txt file
                File.AppendAllText(GlobalVariables.trainingDataFilePath, Environment.NewLine + addNewEmailSubject);

                oClient.Quit();
            }

            
            
            //Completed Loop
            Console.WriteLine("Completed Execution.....");
        }
    }
}
