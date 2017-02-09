using System;
using System.Collections.Generic;
using System.Text;
using System.Data.Common;
using System.Data.OleDb;

namespace DbReader
{
    public class ExcelReader
    {
        #region === Members ===

        string[] m_Fields;
        string m_DbPath;
        string m_Page;

        #endregion

        #region === Constractors ===

        public ExcelReader()
        {
            m_Fields = null;
            m_DbPath = "";
            m_Page = "";
        }

        public ExcelReader(string DbPath, string[] Fields, string Page)
        {
            m_Fields = Fields;
            m_DbPath = DbPath;
            m_Page = Page;
        }

        #endregion

        #region === Public ===

        public string[] Fields
        {
            get { return m_Fields; }
            set { m_Fields = value; }
        }

        public string DbPath
        {
            get { return m_DbPath; }
            set { m_DbPath = value; }
        }

        public string Page
        {
            get { return m_Page; }
            set { m_Page = value; }
        }

        public int[,] GetIntegerTable()
        {
            int Columns = m_Fields.Length;
            System.Collections.ArrayList[] table = new System.Collections.ArrayList[Columns];
            int i;
            string value;

            for (i = 0; i < Columns; i++)
            {
                table[i] = new System.Collections.ArrayList();
            }

            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + m_DbPath + @" ; Extended Properties=""Excel 8.0;HDR=YES;""";
            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");
            DbConnection connection = factory.CreateConnection();

            connection.ConnectionString = connectionString;

            DbCommand command = connection.CreateCommand();

            string commandText = "SELECT " + m_Fields[0];

            for (i = 1; i < Columns; i++)
            {
                commandText += ", " + m_Fields[i];
            }

            commandText += " FROM [" + m_Page + "$]";
            command.CommandText = commandText;

            connection.Open();

            DbDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                double A;

                for (i = 0; i < Columns; i++)
                {
                    value = dr[m_Fields[i]].ToString();

                    if (value.Length == 0)
                    {
                        A = 0;
                    }
                    else
                    {
                        if (!double.TryParse(value, out A))
                        {
                            A = 0;
                        }
                    }

                    int nA = Convert.ToInt32(A);

                    table[i].Add(nA);
                }
            }

            connection.Close();

            int Rows = table[0].Count;
            int[,] IntegerTable = new int[Rows, Columns];

            for (i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Columns; j++)
                {
                    IntegerTable[i, j] = (int)table[j][i];
                }
            }

            // release items from table
            for (i = 0; i < Columns; i++)
            {
                table[i].Clear();
            }

            return IntegerTable;
        }

        public double[,] GetDoubleTable()
        {
            int Columns = m_Fields.Length;
            System.Collections.ArrayList[] table = new System.Collections.ArrayList[Columns];
            int i;
            string value;

            for (i = 0; i < Columns; i++)
            {
                table[i] = new System.Collections.ArrayList();
            }

            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + m_DbPath + @" ; Extended Properties=""Excel 8.0;HDR=YES;""";
            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");
            DbConnection connection = factory.CreateConnection();

            connection.ConnectionString = connectionString;

            DbCommand command = connection.CreateCommand();

            string commandText = "SELECT " + m_Fields[0];

            for (i = 1; i < Columns; i++)
            {
                commandText += ", " + m_Fields[i];
            }

            commandText += " FROM [" + m_Page + "$]";
            command.CommandText = commandText;

            connection.Open();

            DbDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                double A;

                for (i = 0; i < Columns; i++)
                {
                    value = dr[m_Fields[i]].ToString();

                    if (value.Length == 0)
                    {
                        A = 0;
                    }
                    else
                    {
                        if (!double.TryParse(value, out A))
                        {
                            A = 0;
                        }
                    }

                    table[i].Add(A);
                }
            }

            connection.Close();

            int Rows = table[0].Count;
            double[,] DoubleTable = new double[Rows, Columns];

            for (i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Columns; j++)
                {
                    DoubleTable[i, j] = (double)table[j][i];
                }
            }

            // release items from table
            for (i = 0; i < Columns; i++)
            {
                table[i].Clear();
            }

            return DoubleTable;
        }

        public string[,] GetStringTable()
        {
            int Columns = m_Fields.Length;
            System.Collections.ArrayList[] table = new System.Collections.ArrayList[Columns];
            int i;

            for (i = 0; i < Columns; i++)
            {
                table[i] = new System.Collections.ArrayList();
            }

            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + m_DbPath + @" ; Extended Properties=""Excel 8.0;HDR=YES;""";
            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");
            DbConnection connection = factory.CreateConnection();

            connection.ConnectionString = connectionString;

            DbCommand command = connection.CreateCommand();

            string commandText = "SELECT " + m_Fields[0];

            for (i = 1; i < Columns; i++)
            {
                commandText += ", " + m_Fields[i];
            }

            commandText += " FROM [" + m_Page + "$]";
            command.CommandText = commandText;

            connection.Open();

            DbDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                string A;

                for (i = 0; i < Columns; i++)
                {
                    A = dr[m_Fields[i]].ToString();
                    table[i].Add(A);
                }
            }

            connection.Close();

            int Rows = table[0].Count;
            string[,] StringTable = new string[Rows, Columns];

            for (i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Columns; j++)
                {
                    StringTable[i, j] = (string)table[j][i];
                }
            }

            // release items from table
            for (i = 0; i < Columns; i++)
            {
                table[i].Clear();
            }

            return StringTable;
        }


        #endregion
    }
}
