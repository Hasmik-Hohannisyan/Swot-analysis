using System;
using System.IO;
using System.Linq;
using System.Data.OleDb;
using SWOT.Del;

namespace SWOT
{
    internal delegate void Gnahatakanner(double[,] arr2);
    internal delegate void GitakutyanAstijan(double[,] arr2, double[] d);
    internal delegate void Kshirner(double[,] arr2, double[] d, double[] w);
    internal delegate void artacelKshirnery(double[] arr1);

    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            //n
            Gnahatakanner gnDel;
            Console.Write("Ներմուծել փորձագետների քանակը՝ ");
            Pordzaget.Count = Convert.ToInt32(Console.ReadLine());
            double[,] PnGn = new double[Pordzaget.Count, Pordzaget.Count];

            do
            {
                gnDel = Pordzaget.MimiancGnahatum;
                gnDel.Invoke(PnGn);
            }
            while (PnGn[PnGn.GetLength(0) - 1, PnGn.GetLength(1) - 1] == -1);

            for (int swot_i = 0; swot_i < 4; swot_i++)
            {
                bool IS_S_W_O_T;
                char S_W_O_T;
                do
                {
                    IS_S_W_O_T = true;
                    Console.Write("S_W_O_T --> ");
                    S_W_O_T = Convert.ToChar(Console.ReadLine());
                    S_W_O_T = char.ToUpper(S_W_O_T);
                    try
                    {
                        if (S_W_O_T != 'S' && S_W_O_T != 'W' && S_W_O_T != 'O' && S_W_O_T != 'T')
                        {
                            IS_S_W_O_T = false;
                            throw new InputException("SWOT վերլուծություն կատարելու համար օգտագործեք համապատասխան տառերը՝ S_W_O_T\t\t");
                        }
                    }
                    catch (InputException InEx)
                    {
                        Console.WriteLine();
                        Console.BackgroundColor = ConsoleColor.DarkRed;
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.WriteLine($"\t\tInputException: {InEx.Message}");
                        Console.BackgroundColor = ConsoleColor.Black;
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine();
                    }
                } while (!IS_S_W_O_T);
                //m
                Gorconner gorc = new Gorconner();
                double[,] gn = new double[Pordzaget.Count, gorc.count];

                Console.WriteLine();
                Console.WriteLine("Գնահատել հետևյալ հարցերը`");
                Console.WriteLine();

                switch (S_W_O_T)
                {
                    case 'S':
                        Console.WriteLine(Resource.S);
                        Console.WriteLine();
                        goto default;
                    case 'W':
                        Console.WriteLine(Resource.W);
                        Console.WriteLine();
                        goto default;
                    case 'O':
                        Console.WriteLine(Resource.O);
                        Console.WriteLine();
                        goto default;
                    case 'T':
                        Console.WriteLine(Resource.T);
                        Console.WriteLine();
                        goto default;
                    default:
                        {
                            Console.WriteLine("Գնահատականները պետք է տրվեն [1-10]-ում։");
                            Console.WriteLine();

                            gnDel = gorc.gnahatakanner;
                            gnDel(gn);

                            gnDel = gorc.chapakargvacGnahatakanner;
                            gnDel(gn);

                            double[] d = new double[Pordzaget.Count];

                            GitakutyanAstijan gitAst = Pordzaget.GitakutyanAstijan;
                            gitAst(PnGn, d);
                            Console.WriteLine();

                            double[] w = new double[gorc.count];
                            Kshirner ksh = gorc.Kshirner;
                            ksh(gn, d, w);
                            Console.WriteLine();
                            artacelKshirnery artKsh = gorc.artacelKshirnery;
                            artKsh(w);

                            for (int i = 0; i < w.GetLength(0); i++)
                            {
                                if (w[i] == w.Max())
                                {
                                    Console.WriteLine();
                                    Console.BackgroundColor = ConsoleColor.DarkYellow;
                                    Console.ForegroundColor = ConsoleColor.Black;
                                    Console.WriteLine($"\t\tW({i + 1}) - ը ունի ամենաբարձր կշռային գործակիցը. ({w[i]})\t\t");
                                    Console.BackgroundColor = ConsoleColor.Black;
                                    Console.ForegroundColor = ConsoleColor.White;
                                }
                            }

                            OleDbConnectionStringBuilder conString = new OleDbConnectionStringBuilder();
                            conString.DataSource = @"C:\Users\E540\OneDrive\Рабочий стол\Verlucutyun.accdb";
                            if (!File.Exists(conString.DataSource))
                            {
                                Console.WriteLine($"{conString.DataSource} հղումը կամ ֆայլը չի գտնվել։");
                                return;
                            }

                            conString.Provider = "Microsoft.ACE.OLEDB.12.0";
                            OleDbConnection con = new OleDbConnection(conString.ConnectionString);

                            try
                            {
                                con.Open();

                                OleDbCommand cmd = new OleDbCommand("", con);
                                cmd.CommandType = System.Data.CommandType.Text;
                                cmd.CommandText = "INSERT INTO Ardyunqner([Name], [W1], [W2], [W3], [W4], [W5], [W6], [MAX]) VALUES(@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8)";
                                cmd.Parameters.Add(new OleDbParameter("@p1", S_W_O_T));
                                cmd.Parameters.Add(new OleDbParameter("@p2", w[0]));
                                cmd.Parameters.Add(new OleDbParameter("@p3", w[1]));
                                cmd.Parameters.Add(new OleDbParameter("@p4", w[2]));
                                cmd.Parameters.Add(new OleDbParameter("@p5", w[3]));
                                cmd.Parameters.Add(new OleDbParameter("@p6", w[4]));
                                cmd.Parameters.Add(new OleDbParameter("@p7", w[5]));
                                for (int i = 0; i < w.GetLength(0); i++)
                                {
                                    if (w[i] == w.Max())
                                    {
                                        cmd.Parameters.Add(new OleDbParameter("@p8", $"w{i + 1}"));
                                    }
                                }
                                cmd.ExecuteNonQuery();
                            }
                            catch (OleDbException ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            finally
                            {
                                if (con != null && con.State == System.Data.ConnectionState.Open)
                                {
                                    con.Close();
                                }
                            }
                            break;
                        }
                }
            }
        }
    }


    internal static class Pordzaget
    {
        private static int count;

        internal static int Count { get { return count; } set { if (value > 1) { count = value; } else { count = 2; } } }

        internal static void MimiancGnahatum(double[,] pnGn)
        {
            try
            {
                for (int i = 0; i < pnGn.GetLength(0); i++)
                {
                    Console.WriteLine($"\t\tԳնահատում է {i + 1} փորձագետը. ");
                    for (int j = 0; j < pnGn.GetLength(1); j++)
                    {
                        Console.Write($"{j + 1} փորձագետին: ");
                        pnGn[i, j] = double.Parse(Console.ReadLine());
                        if (pnGn[i, j] > 1 || pnGn[i, j] < 0)
                        {
                            pnGn[pnGn.GetLength(0) - 1, pnGn.GetLength(1) - 1] = -1;
                            Console.WriteLine();
                            throw (new NumbeException("Փորձագետները միմյանց մասնակցությունը գնահատում են 0 կամ 1 թվանշաններով\t\t\n\t" +
                                  "\t\t\t1 - դրական կարծիք, 0 - բացասական կարծիք\t\t\t\t\t"));
                        }
                    }
                }
            }
            catch (NumbeException ex)
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.ForegroundColor = ConsoleColor.Black;
                Console.WriteLine($"\tNumber exception: { ex.Message}");
                Console.BackgroundColor = ConsoleColor.Black;
                Console.ForegroundColor = ConsoleColor.White;
            }
            Console.WriteLine();
        }

        internal static void GitakutyanAstijan(double[,] pnGn, double[] d)
        {
            double yndhanur_1_qanak = 0;
            for (int i = 0; i < pnGn.GetLength(0); i++)
            {
                double toxi_1_qanak = 0;
                for (int j = 0; j < pnGn.GetLength(1); j++)
                {
                    if (pnGn[i, j] == 1)
                    {
                        toxi_1_qanak++;
                    }
                }
                yndhanur_1_qanak += toxi_1_qanak;
                d[i] = toxi_1_qanak;
            }
            for (int j = 0; j < pnGn.GetLength(1); j++)
            {
                d[j] /= yndhanur_1_qanak;
            }
        }
    }

    internal class Gorconner
    {
        internal readonly int count = 6;
        private string path;

        internal string Path { set { this.path = value; } }
       
        internal void gnahatakanner(double[,] gn)
        {
            for (int i = 0; i < gn.GetLength(0); i++)
            {
                Console.WriteLine($"\t\tԳնահատում է {i + 1} փորձագետը: ");
                for (int j = 0; j < gn.GetLength(1); j++)
                {
                    Console.Write($"{j + 1} օբյեկտի գնահատական: ");
                    gn[i, j] = double.Parse(Console.ReadLine());
                }
                Console.WriteLine();
            }
        }

        internal void chapakargvacGnahatakanner(double[,] gn)
        {
            for (int i = 0; i < gn.GetLength(0); i++)
            {
                double gumar = 0;
                for (int j = 0; j < gn.GetLength(1); j++)
                {
                    gumar += gn[i, j];
                }
                for (int j = 0; j < gn.GetLength(1); j++)
                {
                    gn[i, j] /= gumar;
                }
            }
        }

        internal void Kshirner(double[,] gn, double[] d, double[] w)
        {
            for (int j = 0; j < gn.GetLength(1); j++)
            {
                for (int i = 0; i < gn.GetLength(0); i++)
                {
                    w[j] += gn[i, j] * d[i];
                }
            }
        }

        internal void artacelKshirnery(double[] w)
        {
            for (int i = 0; i < w.GetLength(0); i++)
            {
                Console.WriteLine($"W({i + 1}) = {w[i]}");
            }
        }
    }

    internal class NumbeException : Exception
    {
        public NumbeException(string message) : base(message)
        {
        }
    }

    internal class InputException : Exception
    {
        public InputException(string message) : base(message)
        {
        }

    }
}