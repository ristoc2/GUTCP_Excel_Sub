using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//Directions:  Build as C# console project.  Place executable in directory with Brilliant Light Power 1-20 Electron Atoms
//spreadsheets and execute (double click or invoke without parameters at command line). Text files for each spreadsheet
//will be generated in which each calculated ionization energy is reduced to one formula by substituting cell formulas for 
//all references to them, substituting physical constants when they are recognized.  Calculated energy has a symbol version
//(suitable for pasting into wxMaxima inside tex() to generate Latex viewable for instance in Texworks, and also for 
//comparison to Brilliant Light's book) and a number version (suitable for pasting into wxMaxima to check against the 
//spreadsheet and compare to measured). Tested with Brilliant Light Power files believed to be current as of 5/30/2018.  
//Table start and column indices are hardcoded per spreadsheet and may need to change if spreadsheets are updated.
namespace FormulaPlugger
{
    class FormulaPlugger
    {
        static void Main(string[] args)
        {

            Application app;
            Workbook book;
            Worksheet sheet;
            Range rng;
            //string formulaSymbolsStr = "";
            string[] nameSplit = new string[5000];
            string[] formulaSplit = new string[5000];
            string[] valueSplit = new string[5000];
            string[] formulaSymbolsSplit = new string[5000];
            string[] formulaSymbolsSplitLatex = new string[5000];
            string fileToOpen = "";
            string fileRel = "";
            int numElectrons = 1;

            app = null;
            IEnumerable<string> files = System.IO.Directory.EnumerateFiles(".");
            for (numElectrons = 1; numElectrons <= 20; numElectrons++)
            {
                foreach (string file in files)
                {
                    if (file.Contains(".xls"))
                    {
                        if (numElectrons == 1 && file.Contains("1e atoms SR NIST.xls"))
                        {
                            System.Console.WriteLine($"Found file {file}");
                            fileRel = file;
                            break;
                        }
                        else if (file.Contains("atoms spreadsheet.xls"))
                        {
                            if (file.Contains($"\\{numElectrons} e"))
                            {
                                System.Console.WriteLine($"Found file {file} with electrons {numElectrons}");
                                fileRel = file;
                                break;
                            }
                        }
                    }
                }
                fileToOpen = System.IO.Path.GetFullPath(fileRel);
                System.Console.WriteLine(fileToOpen);
                app = new Application();
                book = null;
                book = app.Workbooks.Open(fileToOpen);
                string toWriteString = "";
                try
                {
                    sheet = (Worksheet)book.Worksheets.get_Item(1);
                    rng = sheet.UsedRange;
                    string formulaStr = "";
                    string cellNameStr = "";
                    string valueStr = "";
                    int formulaCnt = 0;
                    foreach (Range c in rng.Cells)
                    {
                        string formula = c.Formula;
                        string cellName = c.Address;
                        formulaStr += formula + ";";
                        cellNameStr += cellName + ";";
                        valueStr += c.Value2 + ";";
                        formulaCnt++;
                    }

                    //System.Console.WriteLine(formulaStr);
                    //System.Console.WriteLine("\n\n");
                    cellNameStr = cellNameStr.Replace("$", "");
                    formulaStr = formulaStr.Replace("$", "");
                    formulaStr = formulaStr.Replace("=", "");
                    formulaStr = formulaStr.Replace("PI()", "3.141592654");
                    formulaStr = formulaStr.Replace("SQRT", "sqrt");
                    formulaStr = formulaStr.Replace("COS", "cos");
                    formulaStr = formulaStr.Replace("SIN", "sin");
                    //formulaStr = formulaStr.Replace("^", "**");
                    //System.Console.WriteLine(cellNameStr);
                    //System.Console.WriteLine(formulaStr);
                    valueSplit = valueStr.Split(';');
                    formulaSplit = formulaStr.Split(';');
                    nameSplit = cellNameStr.Split(';');
                    formulaSymbolsSplit = formulaStr.Split(';');
                    //formulaSymbolsSplitLatex = formulaStr.Split(';');
                    int foundRef = 1;
                    string toReplace;
                    string toReplaceWith;
                    while (foundRef == 1)
                    {
                        foundRef = 0;
                        for (int i = 0; i < formulaCnt; i++)
                        {
                            while (HasCellRef(formulaSplit[i]) != -1)
                            {
                                foundRef = 1;
                                toReplace = GetFirstCellRef(formulaSplit[i]);
                                for (int j = 0; j < formulaCnt; j++)
                                {
                                    if (StrMatch(nameSplit[j].ToString(), toReplace.ToString()))
                                    {
                                        toReplaceWith = "(" + formulaSplit[j] + ")"; // TODO add parentheses
                                        formulaSplit[i] = formulaSplit[i].Replace(nameSplit[j], toReplaceWith);


                                        //System.Console.WriteLine($"Replacing {nameSplit[j]} with {toReplaceWith}");
                                    }
                                }
                            }
                        }
                    }
                    int ZCol = 0;
                    int totalCols = 1;
                    int calcCol = 0;
                    int measCol = 0;
                    int errCol = 0;
                    int startCell = 0;
                    if (numElectrons == 1)
                    {
                        startCell = 88;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 6;
                        measCol = 7;
                        errCol = 9;
                    }
                    else if (numElectrons == 2)
                    {
                        startCell = 252;
                        totalCols = 21;
                        ZCol = 0;
                        calcCol = 10;
                        measCol = 12;
                        errCol = 14;
                    }
                    else if (numElectrons == 3)
                    {
                        startCell = 1886;
                        totalCols = 23;
                        ZCol = 0;
                        calcCol = 15;
                        measCol = 16;
                        errCol = 18;
                    }
                    else if (numElectrons == 4)
                    {
                        startCell = 19;
                        totalCols = 19;
                        ZCol = 0;
                        calcCol = 16;
                        measCol = 17;
                        errCol = 18;
                    }
                    else if (numElectrons == 5)
                    {
                        startCell = 20;
                        totalCols = 20;
                        ZCol = 0;
                        calcCol = 15;
                        measCol = 16;
                        errCol = 17;
                    }
                    else if (numElectrons == 6)
                    {
                        startCell = 18;
                        totalCols = 18;
                        ZCol = 0;
                        calcCol = 15;
                        measCol = 16;
                        errCol = 17;
                    }
                    else if (numElectrons == 7)
                    {
                        startCell = 18;
                        totalCols = 18;
                        ZCol = 0;
                        calcCol = 15;
                        measCol = 16;
                        errCol = 17;
                    }
                    else if (numElectrons == 8)
                    {
                        startCell = 18;
                        totalCols = 18;
                        ZCol = 0;
                        calcCol = 15;
                        measCol = 16;
                        errCol = 17;
                    }
                    else if (numElectrons == 9)
                    {
                        startCell = 18;
                        totalCols = 18;
                        ZCol = 0;
                        calcCol = 15;
                        measCol = 16;
                        errCol = 17;
                    }
                    else if (numElectrons == 10)
                    {
                        startCell = 18;
                        totalCols = 18;
                        ZCol = 0;
                        calcCol = 15;
                        measCol = 16;
                        errCol = 17;
                    }
                    else if (numElectrons == 11)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 12)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 13)
                    {
                        startCell = 17;
                        totalCols = 17;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 14)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 15)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 16)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 17)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 18)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 19)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    else if (numElectrons == 20)
                    {
                        startCell = 11;
                        totalCols = 11;
                        ZCol = 0;
                        calcCol = 8;
                        measCol = 9;
                        errCol = 10;
                    }
                    int printing = 0;
                    var varNameDict = GetVariableName();
                    var varNameDictLatex = GetVariableNameLatex();
                    for (int i = 0; i < formulaCnt; i++)
                    {
                        //System.Console.WriteLine(i);
                        //System.Console.WriteLine(formulaSplit[i]);
                        if (formulaSplit[i] != "" && i % totalCols == 0)
                        {
                            if (i == startCell)
                            {
                                printing = 1;
                                System.Console.WriteLine($"{numElectrons} electron atoms:");
                                toWriteString += $"{ numElectrons} electron atoms:\r\n";
                            }
                        }
                        else if (formulaSplit[i] == "" && i % totalCols == 0)
                        {
                            printing = 0;
                        }

                        if (i % totalCols == ZCol)
                        {
                            if (printing == 1)
                            {
                                System.Console.Write(GetElement(Int32.Parse(valueSplit[i])));
                                toWriteString += GetElement(Int32.Parse(valueSplit[i]));
                                if (numElectrons > 1)
                                {
                                    System.Console.WriteLine("");
                                    toWriteString += "\r\n";
                                }
                            }
                        }
                        if (i % totalCols == 1 && numElectrons == 1)
                        {
                            if (printing == 1)
                            {
                                System.Console.Write("-");
                                toWriteString += "-";
                                System.Console.WriteLine(valueSplit[i]);
                                toWriteString += $"{valueSplit[i]}\r\n";
                            }
                        }
                        if (i % totalCols == calcCol)
                        {
                            if (printing == 1)
                            {
                                formulaSymbolsSplit[i] = formulaSplit[i];
                                formulaSymbolsSplitLatex[i] = formulaSplit[i];
                                foreach (KeyValuePair<string, string> kvp in varNameDict)
                                {
                                    formulaSymbolsSplit[i] = formulaSymbolsSplit[i].Replace(kvp.Key, kvp.Value);
                                }
                                //foreach (KeyValuePair<string, string> kvp in varNameDictLatex)
                                //{
                                //    formulaSymbolsSplitLatex[i] = formulaSymbolsSplitLatex[i].Replace(kvp.Key, kvp.Value);
                                //}
                                //System.Console.WriteLine(formulaSymbolsSplitLatex[i]);
                                System.Console.WriteLine($"{numElectrons} electrons, Calculated - formula with variables:");
                                toWriteString += $"{numElectrons} electrons, Calculated - formula with variables:\r\n";
                                System.Console.WriteLine(formulaSymbolsSplit[i]);
                                toWriteString += $"{formulaSymbolsSplit[i]}\r\n";
                                System.Console.WriteLine($"{numElectrons} electrons, Calculated - formula with numbers:");
                                toWriteString += $"{numElectrons} electrons, Calculated - formula with numbers:\r\n";
                                System.Console.WriteLine(formulaSplit[i]);
                                toWriteString += $"{formulaSplit[i]}\r\n";
                                System.Console.WriteLine($"{numElectrons} electrons, Calculated - answer:");
                                toWriteString += $"{numElectrons} electrons, Calculated - answer:\r\n";
                                System.Console.WriteLine(valueSplit[i]);
                                toWriteString += $"{valueSplit[i]}";
                            }
                        }
                        if (i % totalCols == measCol)
                        {
                            if (printing == 1)
                            {
                                System.Console.WriteLine("Measured - answer:");
                                toWriteString += $"Measured - answer:\r\n";
                                System.Console.WriteLine(valueSplit[i]);
                                toWriteString += $"{valueSplit[i]}\r\n";
                            }
                        }
                        if (i % totalCols == errCol)
                        {
                            if (printing == 1)
                            {
                                System.Console.WriteLine("Relative error:");
                                toWriteString += $"Relative error:\r\n";
                                System.Console.WriteLine(valueSplit[i]);
                                toWriteString += $"{valueSplit[i]}\r\n";
                            }
                        }
                    }
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(fileToOpen.Replace(".xls", ".txt")))
                    {
                        sw.WriteLine(toWriteString);
                    }
                }
                finally
                {
                    app.DisplayAlerts = false;
                    app.Quit();
                    Marshal.FinalReleaseComObject(app);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }// end of pasted block
                
                
            }
            
        }

        public static string GetElement(int Z)
        {
            string[] elemTable = { "",
            "H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al",
            "Si", "P", "S", "Cl", "Ar", "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe",
            "Co", "Ni", "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr", "Rb", "Sr", "Y",
            "Zr", "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te",
            "I", "Xe", "Cs", "Ba", "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb",
            "Dy", "Ho", "Er", "Tm", "Yb", "Lu", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt",
            "Au", "Hg", "Tl", "Pb", "Bi", "Po", "At", "Rn", "Fr", "Ra", "Ac", "Th", "Pa",
            "U", "Np", "Pu", "Am", "Cm", "Bk", "Cf", "Es", "Fm", "Md", "No", "Lr", "Rf",
            "Db", "Sg", "Bh", "Hs", "Mt", "Ds", "Rg", "Cn"
            };
            return elemTable[Z];
            
        }
        public static Dictionary<string, string> GetVariableName()
        {
            Dictionary<string, string> varName = new Dictionary<string, string>();
            varName.Clear();
            varName.Add("1.2566371E-06", "nu_0");
            varName.Add("1.60217653E-19", "e");
            varName.Add("1.05457168E-34", "hbar");
            varName.Add("9.1093826E-31", "m_e");
            varName.Add("5.291772108E-11", "a_0");
            varName.Add("2.0023193043718", "g");
            varName.Add("0.007297352568", "alpha");
            varName.Add("8.6602540E-01", "(3/4)^(1/2)");
            varName.Add("299792458", "c");
            varName.Add("5.29465409718E-11", "aH");
            varName.Add("1.67262171E-27", "m_p");
            varName.Add("8.854187817E-12", "epsilon_0");
            varName.Add("3.141592654", "pi");

            varName.Add("1.602189246E-19", "e"); // 3e atom sheet has these slightly different constants
            varName.Add("1.054588757E-34", "hbar");
            varName.Add("9.10953447E-31", "m_e");
            varName.Add("8.854187827E-12", "epsilon_0");
            varName.Add("5.291770644E-11", "a_0");
            return varName;
        }
        
        public static Dictionary<string, string> GetVariableNameLatex()
        {
            Dictionary<string, string> varNameLatex = new Dictionary<string, string>();
            varNameLatex.Clear();
            varNameLatex.Add("1.2566371E-06", "\\nu_0");
            varNameLatex.Add("1.60217653E-19", "e");
            varNameLatex.Add("1.05457168E-34", "\\hbar");
            varNameLatex.Add("9.1093826E-31", "\\m_e");
            varNameLatex.Add("5.291772108E-11", "\\a_0");
            varNameLatex.Add("2.0023193043718", "g");
            varNameLatex.Add("0.007297352568", "\\alpha");
            varNameLatex.Add("8.6602540E-01", "(3/4)^{1/2}");
            varNameLatex.Add("299792458", "c");
            varNameLatex.Add("5.29465409718E-11", "aH");
            varNameLatex.Add("1.67262171E-27", "\\m_p");
            varNameLatex.Add("8.854187817E-12", "\\epsilon_0");
            varNameLatex.Add("3.141592654", "\\pi");
            return varNameLatex;
        }


        public static int IsCapital(string input)
        {
            int i = 0;

            if (input[i] >= 'A' && input[i] <= 'Z')
            {
                return 1;
            }
            return -1;
        }
        public static int HasCellRef(string input)
        {
            int i = 0;
            if (input.Length < 2)
            {
                return -1;
            }
            for (i = 0; i < input.Length - 1; i++)
            {
                if (input[i] >= 'A' && input[i] <= 'Z' && input[i + 1] >= '0' && input[i + 1] <= '9')
                {
                    return i;
                }
            }
            return -1;
        }
        public static string GetFirstCellRef(string input)
        {
            int start = 0;
            int length = 0;

            start = HasCellRef(input);
            if (start == -1)
            {
                return "";
            }
            while (start < input.Length - 1 && !(input[start] >= 'A' && input[start] <= 'Z'))
            {
                start++;
            }
            if (start == input.Length - 1)
            {
                return "";
            }
            length = 1;
            while (start + length < input.Length && input[start + length] >= '0' && input[start + length] <= '9')
            {
                length++;
            }
            return input.Substring(start, length);
        }
        public static bool StrMatch(string str1, string str2)
        {
            int i = 0;
            int len = str1.Length;
            if (len > str2.Length)
            {
                len = str2.Length;
            }
            for (i = 0; i < len; i++)
            {
                if (str1[i] != str2[i])
                {
                    return false;
                }
            }
            if (str1.Length > str2.Length) //matched only a substring such as "A10" and "A101"
            {
                if (str1[i] >= '0' && str1[i] <= '9')
                {
                    return false;
                }
            }
            else if (str1.Length < str2.Length)
            {
                if (str2[i] >= '0' && str2[i] <= '9')
                {
                    return false;
                }
            }
            return true;
        }

    }
}
