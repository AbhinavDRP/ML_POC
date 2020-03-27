namespace ML_Based_Invoice_Prediction
{
    /// <summary>
    /// Main Entry Point of the Program
    /// </summary>
    class Program
    {
        /// <summary>
        /// The entry point of the program, where the program control starts and ends.
        /// </summary>
        /// <param name="args">The command-line arguments.</param>
        static void Main(string[] args)
        {
            ML_Model classObj = new ML_Model();
            classObj.Execute_ML_Model();
        }
    }
}
 