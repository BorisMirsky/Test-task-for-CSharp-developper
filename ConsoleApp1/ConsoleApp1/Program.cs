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

        static Dictionary<string, object> dict_result = new Dictionary<string, object>();

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
            int counter = 0;        // сквозной счетчик
            for (int i = 1; i < (h + 1); i++)
            {   
                for (int j = 0; j < w; j++)
                {
                    string k = alphabet[j] + i.ToString();   // key for dictionary
                    
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

            // Вычисление всех выражений
            var expressionCells = dict_result
                .Where(kv => kv.Value is string s && s.StartsWith("="))
                .Select(kv => kv.Key)
                .ToList();
            foreach (var cell in expressionCells)
            {
                EvaluateCell(cell);
            }

            foreach (KeyValuePair<string, object> entry in dict_result)
            {
                Debug.WriteLine($"Key = {entry.Key}, Value = {entry.Value}");
            }
        }


        // Рекурсивное вычисление значения ячейки
        static int EvaluateCell(string cell)
        {
            object value = dict_result[cell];
            if (value is int)
                return (int)value;
            if (value is string s)
            {
                if (s.StartsWith("="))
                {
                    string expr = s.Substring(1);
                    var (operands, ops) = ParseExpression(expr);
                    int result = GetOperandValue(operands[0]);
                    for (int i = 0; i < ops.Count; i++)
                    {
                        int next = GetOperandValue(operands[i + 1]);
                        switch (ops[i])
                        {
                            case '+': result += next; break;
                            case '-': result -= next; break;
                            case '*': result *= next; break;
                            case '/': result /= next; break;
                        }
                    }
                    dict_result[cell] = result; // сохраняем результат
                    return result;
                }
                else
                {
                    throw new InvalidOperationException($"Cell {cell} contains a string, cannot be used in expression");
                }
            }
            throw new InvalidOperationException($"Cell {cell} has invalid type");
        }

        // Получение числового значения операнда (числа или ссылки)
        static int GetOperandValue(object operand)
        {
            if (operand is int num)
                return num;
            else if (operand is string refCell)
                return EvaluateCell(refCell);
            else
                throw new InvalidOperationException("Invalid operand");
        }


        // Парсинг выражения в список операндов и операторов
        static (List<object> operands, List<char> ops) ParseExpression(string expr)
        {
            List<object> operands = new List<object>();
            List<char> ops = new List<char>();
            int i = 0;
            int len = expr.Length;
            HashSet<char> opSet = new HashSet<char> { '+', '-', '*', '/' };

            while (i < len)
            {
                char c = expr[i];
                if (char.IsDigit(c))
                {
                    int start = i;
                    while (i < len && char.IsDigit(expr[i]))
                        i++;
                    int num = int.Parse(expr.Substring(start, i - start));
                    operands.Add(num);
                    if (i < len && !opSet.Contains(expr[i]))
                        ops.Add('*'); // неявное умножение
                }
                else if (char.IsLetter(c) && c >= 'A' && c <= 'Z')
                {
                    char letter = c;
                    i++;
                    int start = i;
                    while (i < len && char.IsDigit(expr[i]))
                        i++;
                    string cellRef = letter + expr.Substring(start, i - start);
                    operands.Add(cellRef);
                    if (i < len && !opSet.Contains(expr[i]))
                        ops.Add('*'); // неявное умножение
                }
                else if (opSet.Contains(c))
                {
                    ops.Add(c);
                    i++;
                }
                else
                {
                    // Неизвестный символ (не должен встречаться) – пропускаем
                    i++;
                }
            }
            return (operands, ops);
        }
    }
}

