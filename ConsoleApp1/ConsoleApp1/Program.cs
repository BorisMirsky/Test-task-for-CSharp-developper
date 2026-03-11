using System;
using System.Diagnostics;
using System.Diagnostics.Metrics;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] splitted = args;
            int h = 0; 
            try
            {
                h = int.Parse(splitted[0]);
            }
            catch (FormatException)
            {
                Console.WriteLine("The string is not in a valid format.");
            }
            int w = 0; 
            try
            {
                w = int.Parse(splitted[1]);
            }
            catch (FormatException)
            {
                Console.WriteLine("The string is not in a valid format.");
            }

            string[] alphabet = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z".Split(",")[0..w];
            Dictionary<string, object> dict_result = new Dictionary<string, object>();
            int counter = 0;     
            for (int i = 1; i < (h + 1); i++)
            {
                for (int j = 0; j < w; j++)
                {
                    string k = alphabet[j] + i.ToString();
                    dict_result[k] = splitted[2..w][counter];
                    counter += 1;
                }
            }
            Debug.WriteLine(dict_result);
        }
    }
}

