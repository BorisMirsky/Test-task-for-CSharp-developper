using System;
using System.Diagnostics;
using System.Diagnostics.Metrics;



namespace ConsoleApp1
{
    class Program
    {

        static void Main(string[] args)
        {
            args = ["3", "4", "12", "=C2", "3", "`Sample", "=A1+B1C1/5", "=A2B1", "=B3-C3", "`Spread", "`Test", "=4-3", "5", "`Sheet"];
            MyMethod(args); 
        }


        static void MyMethod(string[] args)
        {

            int h = 0;
            try
            {
                h = int.Parse(args[0]);
            }
            catch (FormatException)
            {
                Console.WriteLine("The string is not in a valid format.");
            }

            int w = 0;
            try
            {
                w = int.Parse(args[1]);
            }
            catch (FormatException)
            {
                Console.WriteLine("The string is not in a valid format.");
            }

            string[] alphabet = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z".Split(",")[0..w];
            Dictionary<string, object> dict_result = new Dictionary<string, object>();
            int counter = 0;                             // сквозной счетчик
            for (int i = 1; i < (h + 1); i++)
            {   
                for (int j = 0; j < w; j++)
                {
                    string k = alphabet[j] + i.ToString();   // key 
                    
                    // '`' string 
                    if (args[2..][counter][0] == '`')
                    {
                        dict_result[k] = (string)args[2..][counter][1..];
                    }

                    // '=' expression
                    else if (args[2..][counter][0] == '=')
                    {
                        dict_result[k] = (string)args[2..][counter];
                    }
                    
                    // int
                    else
                    {
                        try
                        {
                            int result = Int32.Parse(args[2..][counter]);
                            dict_result[k] = result;
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine($"Unable to parse");
                        }
                    }
                        counter += 1;
                }
            }

            foreach (KeyValuePair<string, object> entry in dict_result)
            {
                Debug.WriteLine($"Key: {entry.Key}, Value: {entry.Value}");
            }
        }
    }
}

