using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;

namespace SReport_Utility
{
    public class KitConstant
    {
        static public string Steel_st2 => "Сталь Ст2";
        static public string Steel_st3 => "Сталь Ст3";
        static public string Steel_20 => "Сталь 20";
        static public string Steel_20x13 => "Сталь 20X13";
        static public string Steel_40x => "Сталь 40X";
        static public string Steel_35 => "Сталь 35";
        static public string Steel_45 => "Сталь 45";
        static public string Steel_09G2C => "Сталь 09Г2С";
        static public string Steel_AISI => "AISI"; 
        static public string Steel_Titan => "Titan";
        static public double HydroTest_T1 => 20;
    }

    public class SReport
    {
        protected DataTable m_DataTable = new DataTable("SReport");

        #region === Public ===

        public int Count
        {
            get { return m_DataTable.Rows.Count; }
        }

        public bool Save(string path)
        {
            bool result = true;

            try
            {
                m_DataTable.WriteXml(path);
            }
            catch
            {
                result = false;
            }

            return result;
        }

        public bool Open(string path)
        {
            bool result = true;

            m_DataTable.Rows.Clear();

            try
            {
                m_DataTable.ReadXml(path);
            }
            catch
            {
                result = false;
            }

            return result;
        }

        #endregion

    }

    public class SReportArchive : SReport
    {
        public SReportArchive()
        {
            // schema
            m_DataTable.Columns.Add("PTO_Type", typeof(int));
            m_DataTable.Columns.Add("PTO_PlateCount", typeof(string));
            m_DataTable.Columns.Add("PTO_PlateMaterial", typeof(int));
            m_DataTable.Columns.Add("PTO_PlateThickness", typeof(int));
            m_DataTable.Columns.Add("PTO_WrapperMaterial1", typeof(int));
            m_DataTable.Columns.Add("PTO_WrapperMaterial2", typeof(int));

            m_DataTable.Columns.Add("PTO_Volume", typeof(string));
            m_DataTable.Columns.Add("PTO_Len", typeof(int));
            m_DataTable.Columns.Add("PTO_WeightNetto", typeof(string));
            m_DataTable.Columns.Add("PTO_WeightBrutto", typeof(string));

            m_DataTable.Columns.Add("PTO_PR", typeof(string));
            m_DataTable.Columns.Add("PTO_PG", typeof(string));
            m_DataTable.Columns.Add("PTO_T", typeof(string));
            m_DataTable.Columns.Add("PTO_PlaneNumber", typeof(string));

            m_DataTable.Columns.Add("People_Razrab", typeof(string));
            m_DataTable.Columns.Add("People_Prov", typeof(string));
            m_DataTable.Columns.Add("People_Nach", typeof(string));
            m_DataTable.Columns.Add("People_Norm", typeof(string));
            m_DataTable.Columns.Add("People_Utv", typeof(string));

            m_DataTable.Columns.Add("Medium_MediumHot", typeof(string));
            m_DataTable.Columns.Add("Medium_MediumCold", typeof(string));
            m_DataTable.Columns.Add("Medium_TypeHot", typeof(int));
            m_DataTable.Columns.Add("Medium_TypeCold", typeof(int));
            m_DataTable.Columns.Add("Medium_MaxThot", typeof(string));
            m_DataTable.Columns.Add("Medium_MaxTcold", typeof(string));
            m_DataTable.Columns.Add("Medium_MinT", typeof(string));

            m_DataTable.Columns.Add("Material_StaticPlate", typeof(int));
            m_DataTable.Columns.Add("Material_PriPlate", typeof(int));
            m_DataTable.Columns.Add("Material_NaprUp", typeof(int));
            m_DataTable.Columns.Add("Material_NaprDown", typeof(int));
            m_DataTable.Columns.Add("Material_Rez", typeof(int));
            m_DataTable.Columns.Add("Material_Shplilka", typeof(int));
            m_DataTable.Columns.Add("Material_Bolt", typeof(int));
            m_DataTable.Columns.Add("Material_Gaika", typeof(int));
            m_DataTable.Columns.Add("Material_Phlance", typeof(int));


            // initialize
            DataRow dr = m_DataTable.NewRow();

            dr["PTO_Type"] = 0;
            dr["PTO_PlateCount"] = "22";
            dr["PTO_PlateMaterial"] = 0;
            dr["PTO_PlateThickness"] = 0;
            dr["PTO_WrapperMaterial1"] = 0;
            dr["PTO_WrapperMaterial2"] = 0;

            dr["PTO_Volume"] = "22";
            dr["PTO_Len"] = 0;
            dr["PTO_WeightNetto"] = "22";
            dr["PTO_WeightBrutto"] = "22";

            dr["PTO_PR"] = "22";
            dr["PTO_PG"] = "22";
            dr["PTO_T"] = "22";
            dr["PTO_PlaneNumber"] = "22";

            dr["People_Razrab"] = "22";
            dr["People_Prov"] = "22";
            dr["People_Nach"] = "22";
            dr["People_Norm"] = "22";
            dr["People_Utv"] = "22";

            dr["Medium_MediumHot"] = "22";
            dr["Medium_MediumCold"] = "22";
            dr["Medium_TypeHot"] = 0;
            dr["Medium_TypeCold"] = 0;
            dr["Medium_MaxThot"] = "22";
            dr["Medium_MaxTcold"] = "22";
            dr["Medium_MinT"] = "22";

            dr["Material_StaticPlate"] = 0;
            dr["Material_PriPlate"] = 0;
            dr["Material_NaprUp"] = 0;
            dr["Material_NaprDown"] = 0;
            dr["Material_Rez"] = 0;
            dr["Material_Shplilka"] = 0;
            dr["Material_Bolt"] = 0;
            dr["Material_Gaika"] = 0;
            dr["Material_Phlance"] = 0;

            // add row to table
            m_DataTable.Rows.Add(dr);
        }

        #region === Public ===

        public void SetValue(string field, object Value)
        {
            DataRow dr = m_DataTable.Rows[0];
            dr[field] = Value;
        }

        public int GetIntegerValue(string field)
        {
            DataRow dr = m_DataTable.Rows[0];
            int result = (int)dr[field];

            return result;
        }

        public string GetStringValue(string field)
        {
            DataRow dr = m_DataTable.Rows[0];
            string result = (string)dr[field];

            return result;
        }

        #endregion
    }

    public class SReportDataBase : SReport
    {
        int m_id = 0;
        string m_startupPath = string.Empty;

        public SReportDataBase()
        {
            // schema
            m_DataTable.Columns.Add("BD_Path", typeof(string));
            m_DataTable.Columns.Add("BD_Name", typeof(string));
            m_DataTable.Columns.Add("BD_Comment", typeof(string));
        }

        #region === Public ===

        public int Id
        {
            get { return m_id; }
            set { m_id = value; }
        }

        public string StartupPath
        {
            get { return m_startupPath; }
            set { m_startupPath = value; }
        }

        public string Add(string Name, string Comment)
        {
            // initialize
            DataRow dr = m_DataTable.NewRow();
            string path = m_startupPath + m_id.ToString() + ".xml";

            dr["BD_Name"] = Name;
            dr["BD_Comment"] = Comment;
            dr["BD_Path"] = path;

            // add row to table
            m_DataTable.Rows.Add(dr);

            m_id++;

            return path;
        }

        public DataRow GetDataRow(int index)
        {
            return m_DataTable.Rows[index];
        }

        public void RemoveAt(int index)
        {
            m_DataTable.Rows.RemoveAt(index);
        }        

        #endregion
    }

    public abstract class MyHashTable
    {
        #region === Memebers ===

        protected ArrayList m_key = new ArrayList();

        #endregion

        #region === Public ===

        public int KeyCount { get { return m_key.Count; } }

        #endregion
    }

    public class MyIntegerHashTable : MyHashTable
    {
        #region === Members ===

        private ArrayList m_value = new ArrayList();

        #endregion

        #region === Public ===

        public void AddValue(string Key, int[] value)
        {
            m_key.Add(Key);
            m_value.Add(value);
        }

        public int[] GetValue(string Key)
        {
            int[] value = null;

            for (int i = 0; i < KeyCount; i++)
            {
                if (Key == (string)m_key[i])
                {
                    value = (int[])m_value[i];
                    break;
                }
            }

            return value;
        }

        #endregion
    }

    public class MyStringHashTable : MyHashTable
    {
        #region === Memeber ===

        ArrayList m_value1 = new ArrayList();
        ArrayList m_value2 = new ArrayList();

        #endregion

        #region === Public ===

        public void Add(string key, string value)
        {
            m_key.Add(key);
            m_value1.Add(value);
            m_value2.Add("");
        }

        public void Add(string key, string value1, string value2)
        {
            m_key.Add(key);
            m_value1.Add(value1);
            m_value2.Add(value2);
        }

        public string GetValue1(string key)
        {
            string value = "";

            for (int i = 0; i < m_key.Count; i++)
            {
                if ((string)m_key[i] == key)
                {
                    value = (string)m_value1[i];
                    break;
                }
            }
            
            return value;
        }

        public string GetValue2(string key)
        {
            string value = "";

            for (int i = 0; i < m_key.Count; i++)
            {
                if ((string)m_key[i] == key)
                {
                    value = (string)m_value2[i];
                    break;
                }
            }

            return value;
        }

        #endregion
    }

    public class MyDoubleHashTable : MyHashTable
    {
        #region === Members ===

        ArrayList[] m_value;
        string[] m_columns;

        #endregion

        #region === Constructor ===

        public MyDoubleHashTable(string[] Columns)
        {
            int count = Columns.Length;

            m_columns = Columns;
            m_value = new ArrayList[count];

            for (int i = 0; i < count; i++)
            {
                m_value[i] = new ArrayList();
            }            
        }

        public void AddValue(string Key, double[] Vals)
        {
            m_key.Add(Key);

            for (int i = 0; i < m_columns.Length; i++)
            {
                m_value[i].Add(Vals[i]);
            }
        }

        public double GetValue(string Key, string Column)
        {
            double val = 0;

            for (int i = 0; i < m_key.Count; i++)
            {
                if (Key == (string)m_key[i])
                {
                    for (int j = 0; j < m_columns.Length; j++)
                    {
                        if (Column == m_columns[j])
                        {
                            val = (double)m_value[j][i];
                            break;
                        }
                    }

                    break;
                }
            }

            return val;
        }

        #endregion
    }

    public class DigitalProcess
    {
        static public int MathTrancate(int Val, int sign)
        {
            if (sign <= 0)
            {
                throw new Exception("Sign parameter must be more than 0.");
            }

            int power = 1;

            for (int i = 1; i <= sign; i++)
            {
                power *= 10;
            }

            bool bIncrement = (Val % power >= power / 2) ? true : false;
            int result = Val / power;
            if (bIncrement) result++;
            result *= power;

            return result;
        }

        static public int MathTrancateInc(int Val, int sign)
        {
            if (sign <= 0)
            {
                throw new Exception("Sign parameter must be more than 0.");
            }

            int power = 1;

            for (int i = 1; i <= sign; i++)
            {
                power *= 10;
            }

            int result = Val / power + 1;
            result *= power;

            return result;
        }

        static public string GetDoubleFormat(double Val, int Sign)
        {
            if (Sign < 0)
            {
                throw new Exception("Number of sign is incorrect! Sign < 0.");
            }

            string s;

            if (Sign == 0)
            {
                s = Convert.ToInt32(Val).ToString();
            }
            else
            {
                long power = 1;

                for (int i = 0; i < Sign; i++)
                {
                    power *= 10;
                }

                long lval = Convert.ToInt64(Val * power);
                double dval = Convert.ToDouble(lval) / Convert.ToDouble(power);

                s = dval.ToString();
            }

            return s;
        }

        static public string MathTrancate(double Val, int Sign, bool Formated)
        {
            if (Sign < 0)
            {
                throw new Exception("Number of sign is incorrect! Sign < 0.");
            }

            if (Sign > 12)
            {
                throw new Exception("Number of sign is incorrect! Sign > 12.");
            }

            string s;

            if (Sign == 0)
            {
                s = Convert.ToInt32(Val).ToString();
            }
            else
            {
                int i;
                double t1 = Val;

                for (i = 0; i <= Sign; i++)
                {
                    t1 *= 10;
                }

                int nVal = (int)t1;
                int controlDigit = nVal % 10;
                double result;

                if (controlDigit >= 5)
                {
                    double additor = 1.0;

                    for (i = 0; i < Sign; i++)
                    {
                        additor /= 10;
                    }

                    result = Val + additor;
                }
                else
                {
                    result = Val;
                }

                t1 = result;

                for (i = 0; i < Sign; i++)
                {
                    t1 *= 10;
                }

                nVal = (int)t1;
                result = nVal;

                for (i = 0; i < Sign; i++)
                {
                    result /= 10;
                }

                s = result.ToString();
                //s = GetDoubleFormat(result, Sign);
            }

            return (Formated) ? ConvertToDigitFormat(s, Sign) : s;
        }

        static public string ConvertToDigitFormat(string sVal, int Sign)
        {
            System.Globalization.NumberFormatInfo m_nfi = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat;
            string s = sVal;

            if (Sign > 0)
            {
                int pos = s.IndexOf(m_nfi.NumberDecimalSeparator[0]);
                int length = s.Length;

                if (pos < 0)
                {
                    s += m_nfi.NumberDecimalSeparator[0];
                    pos = s.IndexOf(m_nfi.NumberDecimalSeparator[0]);
                    length++;
                }

                for (int i = length - pos; i <= Sign; i++)
                {
                    s += "0";
                }
            }

            return s;
        }

        static public string MathTrancateToDigital(double Val, int Sign)
        {
            // 0 < Val
            // 1.245E-03 -> 0,001245

            if (Sign < 0)
            {
                throw new Exception("Number of sign is incorrect! Sign < 0.");
            }

            if (Sign > 12)
            {
                throw new Exception("Number of sign is incorrect! Sign > 12.");
            }

            if (Val < 0)
            {
                throw new Exception("Input value must be more than 0.");
            }

            string s;
            int i;

            if (Sign == 0)
            {
                s = Convert.ToInt32(Val).ToString();
            }
            else
            {
                double t1 = Val;

                for (i = 0; i <= Sign; i++)
                {
                    t1 *= 10;
                }

                int nVal = (int)t1;
                int controlDigit = nVal % 10;
                double result;

                if (controlDigit >= 5)
                {
                    double additor = 1.0;

                    for (i = 0; i < Sign; i++)
                    {
                        additor /= 10;
                    }

                    result = Val + additor;
                }
                else
                {
                    result = Val;
                }

                s = Convert.ToInt32(result).ToString();
                s += ",";

                for (i = 0; i < Sign; i++)
                {
                    result *= 10;
                    int lVal = (int)result;
                    controlDigit = lVal % 10;
                    s += controlDigit.ToString();
                }

                // remove last zeros
                while (true)
                {
                    if (s[s.Length - 1] == '0')
                    {
                        s = s.Substring(0, s.Length - 1);
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return s;
        }
    }

    public class PrintReport
    {
        DataTable m_DataTable = new DataTable("SReport");

        public PrintReport()
        {
            System.Collections.ArrayList list = new System.Collections.ArrayList();

            list.Add("PTO");
            list.Add("PTO_Plane_Number");
            list.Add("PTO_Full_Plane_Number");
            list.Add("PTO_Plates");
            list.Add("PTO_Plates_1");
            list.Add("PTO_OriginalName");

            list.Add("Medium_Hot_Name");
            list.Add("Medium_Cold_Name");
            list.Add("Medium_Hot_MaxT" );
            list.Add("Medium_Cold_MaxT");
            list.Add("Medium_MinT");
            list.Add("Medium_WallT");

            list.Add("Weight_Netto");
            list.Add("Weight_Brutto");

            list.Add("Pmax");

            list.Add("People_Developer");
            list.Add("People_Auditor");
            list.Add("People_NachPodrz");
            list.Add("People_Ncontr");
            list.Add("People_Utverdil");

            // table 1
            list.Add("Table1_Period");
            list.Add("Table1_Phrase1");

            // table 2
            list.Add("Table2_Material_Row1");
            list.Add("Table2_Material_Row2");
            list.Add("Table2_Material_Row3");
            list.Add("Table2_Material_Row4");
            list.Add("Table2_Material_Row5");
            list.Add("Table2_Material_Row6");
            list.Add("Table2_Material_Row7");
            list.Add("Table2_Material_Row8");

            list.Add("Table2_R18");

            list.Add("Table2_T1");
            list.Add("Table2_T2");

            // предел текучести
            list.Add("Table2_PredelTech_11");
            list.Add("Table2_PredelTech_12");

            list.Add("Table2_PredelTech_21");
            list.Add("Table2_PredelTech_22");

            list.Add("Table2_PredelTech_31");
            list.Add("Table2_PredelTech_32");

            list.Add("Table2_PredelTech_41");
            list.Add("Table2_PredelTech_42");

            list.Add("Table2_PredelTech_51");
            list.Add("Table2_PredelTech_52");

            list.Add("Table2_PredelTech_61");
            list.Add("Table2_PredelTech_62");

            list.Add("Table2_PredelTech_71");
            list.Add("Table2_PredelTech_72");

            list.Add("Table2_PredelTech_81");
            list.Add("Table2_PredelTech_82");

            // предел прочности
            list.Add("Table2_PredelProch_11");
            list.Add("Table2_PredelProch_12");

            list.Add("Table2_PredelProch_21");
            list.Add("Table2_PredelProch_22");

            list.Add("Table2_PredelProch_31");
            list.Add("Table2_PredelProch_32");

            list.Add("Table2_PredelProch_41");
            list.Add("Table2_PredelProch_42");

            list.Add("Table2_PredelProch_51");
            list.Add("Table2_PredelProch_52");

            list.Add("Table2_PredelProch_61");
            list.Add("Table2_PredelProch_62");

            list.Add("Table2_PredelProch_71");
            list.Add("Table2_PredelProch_72");

            list.Add("Table2_PredelProch_81");
            list.Add("Table2_PredelProch_82");

            // Номинальное допускаемое напряжение (сигма)
            list.Add("Table2_Sigma_11");
            list.Add("Table2_Sigma_12");

            list.Add("Table2_Sigma_21");
            list.Add("Table2_Sigma_22");

            list.Add("Table2_Sigma_31");
            list.Add("Table2_Sigma_32");

            list.Add("Table2_Sigma_41");
            list.Add("Table2_Sigma_42");

            list.Add("Table2_Sigma_51");
            list.Add("Table2_Sigma_52");

            list.Add("Table2_Sigma_61");
            list.Add("Table2_Sigma_62");

            list.Add("Table2_Sigma_71");
            list.Add("Table2_Sigma_72");

            list.Add("Table2_Sigma_81");
            list.Add("Table2_Sigma_82");

            // Модуль упругости Eт 10-5, МПа
            list.Add("Table2_ModuleUpr_11");
            list.Add("Table2_ModuleUpr_12");

            list.Add("Table2_ModuleUpr_21");
            list.Add("Table2_ModuleUpr_22");

            list.Add("Table2_ModuleUpr_31");
            list.Add("Table2_ModuleUpr_32");

            list.Add("Table2_ModuleUpr_41");
            list.Add("Table2_ModuleUpr_42");

            list.Add("Table2_ModuleUpr_51");
            list.Add("Table2_ModuleUpr_52");

            list.Add("Table2_ModuleUpr_61");
            list.Add("Table2_ModuleUpr_62");

            list.Add("Table2_ModuleUpr_71");
            list.Add("Table2_ModuleUpr_72");

            list.Add("Table2_ModuleUpr_81");
            list.Add("Table2_ModuleUpr_82");

            // Коэффициент линейного расширения
            //list.Add("Table2_LinearExt_11");
            list.Add("Table2_LinearExt_12");

            //list.Add("Table2_LinearExt_21");
            list.Add("Table2_LinearExt_22");

            //list.Add("Table2_LinearExt_31");
            list.Add("Table2_LinearExt_32");

            //list.Add("Table2_LinearExt_41");
            list.Add("Table2_LinearExt_42");

            //list.Add("Table2_LinearExt_51");
            list.Add("Table2_LinearExt_52");

            //list.Add("Table2_LinearExt_61");
            list.Add("Table2_LinearExt_62");

            //list.Add("Table2_LinearExt_71");
            list.Add("Table2_LinearExt_72");

            list.Add("Table2_LinearExt_82");

            // Table 3
            list.Add("Table3_P");
            list.Add("Table3_SigmaTH");
            list.Add("Table3_SigmaT");
            list.Add("Table3_PH");

            list.Add("Presuure_HydroTest");
            list.Add("Pressure_Calc");

            // Table 4
            list.Add("Table4_B");

            //"Bm", "B", "Am", "A", "b", "Bsh", "d", "z", "c111", "c112", "c12", "c2", "S1", "S2";            

            list.Add("Table_Bm");
            list.Add("Table_Am");
            list.Add("Table_A");
            list.Add("Table_bb");
            list.Add("Table_Bsh");
            list.Add("Table_d");
            list.Add("Table_z");
            list.Add("Table_c111");
            list.Add("Table_c112");
            list.Add("Table_c12");
            list.Add("Table_c2");
            list.Add("Table_S1");
            list.Add("Table_S2");

            list.Add("Table4_Y");
            list.Add("Table4_Km");
            list.Add("Table4_Ko");
            list.Add("Table4_c_stat");
            list.Add("Table4_c_pri");

            list.Add("Table4_SR1");
            list.Add("Table4_SR2");
            list.Add("Table4_SR12");
            list.Add("Table4_SR22");

            list.Add("Table4_Sigma");
            list.Add("Table4_SigmaH");
            list.Add("Table4_Sigma_pri");
            list.Add("Table4_SigmaH_pri");

            list.Add("Table4_SR11_plus_C");
            list.Add("Table4_SR12_plus_C");
            list.Add("Table4_SR21_plus_C");
            list.Add("Table4_SR22_plus_C");

            // Table 5
            list.Add("Table5_Epsilon");
            list.Add("Table5_Am");
            list.Add("Table5_Bm");
            list.Add("Table5_b");
            list.Add("Table5_Q0");
            list.Add("Table5_Fob");
            list.Add("Table5_P");
            list.Add("Table5_Ph");
            list.Add("Table5_Eta");
            list.Add("Table5_Fp");
            list.Add("Table5_Fd");
            list.Add("Table5_Hi");
            list.Add("Table5_m");
            list.Add("Table5_F2");
            list.Add("Table5_F2_plus");
            list.Add("Table5_Fph");

            list.Add("Table5_F2h");
            list.Add("Table5_F2h_plus");
            list.Add("Table5_z");
            list.Add("Table5_d0");
            list.Add("Table5_ksi");
            list.Add("Table5_MomentK");
            list.Add("Table5_Fow");
            list.Add("Table5_Fw");
            list.Add("Table5_dc");
            list.Add("Table5_SigmaW");
            list.Add("Table5_dw");
            list.Add("Table5_dow");

            // table 6
            list.Add("Table6_d0_down");
            list.Add("Table6_d0_up");
            list.Add("Table6_dc");
            list.Add("Table6_Mk_down");
            list.Add("Table6_Mk_up");
            list.Add("Table6_ksi");
            list.Add("Table6_Fow_down");
            list.Add("Table6_Fow_up");
            list.Add("Table6_SigmaW");
            list.Add("Table6_dw_down");
            list.Add("Table6_dw_up");
            list.Add("Table6_d0w_down");
            list.Add("Table6_d0w_up");

            // table 8, page 16
            list.Add("Table8_b0");
            list.Add("Table8_b");
            list.Add("Table8_delta");
            list.Add("Table8_q0");
            list.Add("Table8_Fob");
            list.Add("Table8_P");
            list.Add("Table8_Ph");
            list.Add("Table8_Dm");
            list.Add("Table8_Fp");
            list.Add("Table8_Fd");
            list.Add("Table8_hi");
            list.Add("Table8_m");
            list.Add("Table8_eta");
            list.Add("Table8_F2");
            list.Add("Table8_F2_plus");
            list.Add("Table8_Fph");
            list.Add("Table8_F2h");
            list.Add("Table8_F2h_plus");
            list.Add("Table8_z");
            list.Add("Table8_d0");
            list.Add("Table8_ksi");
            list.Add("Table8_Mk");
            list.Add("Table8_Fow");
            list.Add("Table8_Fw");

            // table 8, page 17
            list.Add("Table8_dc");
            list.Add("Table8_dw");
            list.Add("Table8_d0w");

            // table 9
            list.Add("Table9_n");
            list.Add("Table9_f");
            list.Add("Table9_h");
            list.Add("Table9_dh");
            list.Add("Table9_epsilon");
            list.Add("Table9_b");
            list.Add("Table9_Epr");
            list.Add("Table9_q");
            list.Add("Table9_A");
            list.Add("Table9_B");
            list.Add("Table9_A1");
            list.Add("Table9_B1");
            list.Add("Table9_R");
            list.Add("Table9_tpt");
            list.Add("Table9_p");

            // table 10
            list.Add("Table10_n"); 
            list.Add("Table10_s");
            list.Add("Table10_l1");
            list.Add("Table10_dT");
            list.Add("Table10_lw");
            list.Add("Table10_d");
            list.Add("Table10_z");
            list.Add("Table10_Ew");
            list.Add("Table10_dw");
            list.Add("Table10_Aw");
            list.Add("Table10_LambdaW");
            list.Add("Table10_Alpha1");
            list.Add("Table10_Alpha2");
            list.Add("Table10_eta");
            list.Add("Table10_Mk");
            list.Add("Table10_Fw");

            list.Add("Table10_Fp_1");
            list.Add("Table10_Fp_2");
            list.Add("Table10_Fp_3");
            list.Add("Table10_Fp_4");

            list.Add("Table10_Ft_1");
            list.Add("Table10_Ft_2");
            list.Add("Table10_Ft_3");
            list.Add("Table10_Ft_4");

            list.Add("Table10_Fph_1");
            list.Add("Table10_Fph_2");
            list.Add("Table10_Fph_3");
            list.Add("Table10_Fph_4");

            list.Add("Table10_Fwp_1");
            list.Add("Table10_Fwp_2");
            list.Add("Table10_Fwp_3");
            list.Add("Table10_Fwp_4");

            list.Add("Table10_SigmaW_Round_1");
            list.Add("Table10_SigmaW_Round_2");
            list.Add("Table10_SigmaW_Round_3");
            list.Add("Table10_SigmaW_Round_4");

            list.Add("Table10_SigmaW_Square_1");
            list.Add("Table10_SigmaW_Square_2");
            list.Add("Table10_SigmaW_Square_3");
            list.Add("Table10_SigmaW_Square_4");

            list.Add("Table10_Tau");

            list.Add("Table10_Sigma4W_Round_1");
            list.Add("Table10_Sigma4W_Round_2");
            list.Add("Table10_Sigma4W_Round_3");
            list.Add("Table10_Sigma4W_Round_4");

            list.Add("Table10_Sigma4W_Square_1");
            list.Add("Table10_Sigma4W_Square_2");
            list.Add("Table10_Sigma4W_Square_3");
            list.Add("Table10_Sigma4W_Square_4");

            // table 11
            list.Add("Table11_Q_1");

            list.Add("Table11_L1_1");
            list.Add("Table11_L2_1");
            list.Add("Table11_L3_1");

            list.Add("Table11_F_1");
            list.Add("Table11_M_1");
            list.Add("Table11_W_1");
            list.Add("Table11_Sigma_1");
            list.Add("Table11_13Sigma_1");

            list.Add("Table11_Q_2");
            list.Add("Table11_F_2");
            list.Add("Table11_M_2");
            list.Add("Table11_W_2");
            list.Add("Table11_Sigma_2");
            list.Add("Table11_13Sigma_2");

            list.Add("Table11_Phrase1");

            list.Add("Table11_Q");
            list.Add("Table11_L");
            list.Add("Table11_t");
            list.Add("Table11_h");
            list.Add("Table11_Sigma");
            list.Add("Table11_Ra");
            list.Add("Table11_Sigma_isgib");

            // table 12
            string sPrefix = "Table12_";
            string sPostfix = "_1314";

            list.Add(sPrefix + "F" + sPostfix);
            list.Add(sPrefix + "A" + sPostfix);
            list.Add(sPrefix + "Sigma" + sPostfix);
            list.Add(sPrefix + "SigmaSquare" + sPostfix);
            list.Add(sPrefix + "Q1" + sPostfix);
            list.Add(sPrefix + "mu" + sPostfix);
            list.Add(sPrefix + "z" + sPostfix);
            list.Add(sPrefix + "Ft" + sPostfix);

            list.Add(sPrefix + "Q" + sPostfix);
            list.Add(sPrefix + "d0" + sPostfix);
            list.Add(sPrefix + "tau" + sPostfix);
            list.Add(sPrefix + "tauSquare" + sPostfix);
            list.Add(sPrefix + "Mk" + sPostfix);
            list.Add(sPrefix + "tauK" + sPostfix);
            list.Add(sPrefix + "Sigma4w" + sPostfix);
            list.Add(sPrefix + "17Sigma" + sPostfix);

            sPostfix = "_1516";

            list.Add(sPrefix + "F" + sPostfix);
            list.Add(sPrefix + "A" + sPostfix);
            list.Add(sPrefix + "Sigma" + sPostfix);
            list.Add(sPrefix + "SigmaSquare" + sPostfix);
            list.Add(sPrefix + "Q1" + sPostfix);
            list.Add(sPrefix + "mu" + sPostfix);
            list.Add(sPrefix + "z" + sPostfix);
            list.Add(sPrefix + "Ft" + sPostfix);

            list.Add(sPrefix + "Q" + sPostfix);
            list.Add(sPrefix + "d0" + sPostfix);
            list.Add(sPrefix + "tau" + sPostfix);
            list.Add(sPrefix + "tauSquare" + sPostfix);
            list.Add(sPrefix + "Mk" + sPostfix);
            list.Add(sPrefix + "tauK" + sPostfix);
            list.Add(sPrefix + "Sigma4w" + sPostfix);
            list.Add(sPrefix + "17Sigma" + sPostfix);

            // Table 13
            list.Add("Table13_F_1");
            list.Add("Table13_F_2");
            list.Add("Table13_F_3");
            list.Add("Table13_F_4");
            list.Add("Table13_F_5");
            list.Add("Table13_F_6");
            list.Add("Table13_F_7");
            list.Add("Table13_F_8");

            list.Add("Table13_d_1");
            list.Add("Table13_d_2");
            list.Add("Table13_d_3");
            list.Add("Table13_d_4");
            list.Add("Table13_d_5");
            list.Add("Table13_d_6");
            list.Add("Table13_d_7");
            list.Add("Table13_d_8");

            list.Add("Table13_h_1");
            list.Add("Table13_h_2");
            list.Add("Table13_h_3");
            list.Add("Table13_h_4");
            list.Add("Table13_h_5");
            list.Add("Table13_h_6");
            list.Add("Table13_h_7");
            list.Add("Table13_h_8");

            list.Add("Table13_K1_1");
            list.Add("Table13_K1_2");
            list.Add("Table13_K1_3");
            list.Add("Table13_K1_4");
            list.Add("Table13_K1_5");
            list.Add("Table13_K1_6");
            list.Add("Table13_K1_7");
            list.Add("Table13_K1_8");

            list.Add("Table13_Km");

            list.Add("Table13_z_1");
            list.Add("Table13_z_2");
            list.Add("Table13_z_3");
            list.Add("Table13_z_4");
            list.Add("Table13_z_5");
            list.Add("Table13_z_6");
            list.Add("Table13_z_7");
            list.Add("Table13_z_8");

            list.Add("Table13_TauP_1");
            list.Add("Table13_TauP_2");
            list.Add("Table13_TauP_3");
            list.Add("Table13_TauP_4");
            list.Add("Table13_TauP_5");
            list.Add("Table13_TauP_6");
            list.Add("Table13_TauP_7");
            list.Add("Table13_TauP_8");

            list.Add("Table13_TauSquareP_1");
            list.Add("Table13_TauSquareP_2");
            list.Add("Table13_TauSquareP_3");
            list.Add("Table13_TauSquareP_4");
            list.Add("Table13_TauSquareP_5");
            list.Add("Table13_TauSquareP_6");
            list.Add("Table13_TauSquareP_7");
            list.Add("Table13_TauSquareP_8");

            // table 14
            list.Add("Table14_P");
            list.Add("Table14_F");
            list.Add("Table14_Apr");
            list.Add("Table14_q");
            list.Add("Table14_L");
            list.Add("Table14_z1");
            list.Add("Table14_z2");
            list.Add("Table14_z3");
            list.Add("Table14_Bpr");

            list.Add("Table14_a");
            list.Add("Table14_M1");
            list.Add("Table14_M2");
            list.Add("Table14_M3");
            list.Add("Table14_b1");
            list.Add("Table14_b2");
            list.Add("Table14_b3");
            list.Add("Table14_h");
            list.Add("Table14_W1");
            list.Add("Table14_W2");
            list.Add("Table14_W3");
            list.Add("Table14_SigmaSquare");
            list.Add("Table14_SigmaIs1");
            list.Add("Table14_SigmaIs2");
            list.Add("Table14_SigmaIs3");
            list.Add("Table14_SSIs");

            // table 15
            list.Add("Table15_A_1");
            list.Add("Table15_A_2");
            list.Add("Table15_Sigma_1");
            list.Add("Table15_Sigma_2");
            list.Add("Table15_SigmaSquare");

            // table 16
            list.Add("Table16_NSigma");
            list.Add("Table16_NN");
            list.Add("Table16_A");
            list.Add("Table16_B");
            list.Add("Table16_t");
            list.Add("Table16_K");
            list.Add("Table16_Sigma4w_1");
            list.Add("Table16_Sigma4w_2");
            list.Add("Table16_Sigma4w_3");
            list.Add("Table16_SigmaA_1");
            list.Add("Table16_SigmaA_2");
            list.Add("Table16_SigmaA_3");
            list.Add("Table16_NSquare_1");
            list.Add("Table16_NSquare_2");
            list.Add("Table16_NSquare_3");
            list.Add("Table16_N1");
            list.Add("Table16_N2");
            list.Add("Table16_N3");
            list.Add("Table16_al");

            list.Add("Table16_N1M");
            list.Add("Table16_BM");

            // ==============================================
            
            foreach(string st1 in list)
            {
                m_DataTable.Columns.Add(st1, typeof(string));
            }

            DataRow dr = m_DataTable.NewRow();

            foreach (string st2 in list)
            {
                dr[st2] = "";
            }

            // bitmaps
            m_DataTable.Columns.Add("Picture1", typeof(Bitmap));
            m_DataTable.Columns.Add("Picture2", typeof(Bitmap));

            m_DataTable.Columns.Add("Picture3", typeof(Bitmap));
            m_DataTable.Columns.Add("Picture4", typeof(Bitmap));
            m_DataTable.Columns.Add("Picture5", typeof(Bitmap));

            m_DataTable.Columns.Add("Picture7", typeof(Bitmap));

            m_DataTable.Columns.Add("Picture8", typeof(Bitmap)); // formula page 25
            m_DataTable.Columns.Add("Picture9", typeof(Bitmap)); // formula page 23
                      
            m_DataTable.Rows.Add(dr);
        }

        public DataTable Table
        {
            get { return m_DataTable; }
        }

        public void SetValue(string Name, string Value)
        {
            DataRow dr = m_DataTable.Rows[0];
            dr[Name] = Value;
        }

        public string GetValue(string Name)
        {
            DataRow dr = m_DataTable.Rows[0];
            string result = (string)dr[Name];

            return result;
        }

        public void SetPicture(string Name, Bitmap bm)
        {
            DataRow dr = m_DataTable.Rows[0];
            dr[Name] = bm;
        }

        public Bitmap GetPicture(string Name)
        {
            DataRow dr = m_DataTable.Rows[0];
            Bitmap result = (Bitmap)dr[Name];

            return result;
        }
    }

    public class SimpleImageProccessing
    {
        static public Bitmap Zoom(Bitmap bm, int width, int height)
        {
            Bitmap result = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            Graphics g = Graphics.FromImage(result);
            g.Clear(Color.White);

            float dZoom;

            if (bm.Width > width || bm.Height > height)
            {
                // zoom
                dZoom = (float)width / (float)bm.Width;

                if ((width - dZoom * width) / 2 < 1)
                {
                    dZoom = (float)height / (float)bm.Height;
                }
            }
            else
            {
                // setup picture in the center
                dZoom = 1.0f;
            }

            float X = (width - dZoom * bm.Width) / 2;
            float Y = (height - dZoom * bm.Height) / 2;
            float CX = dZoom * bm.Width;
            float CY = dZoom * bm.Height;

            g.DrawImage(bm, X, Y, CX, CY);

            return result;
        }
    }

    public class SteelProperty : SReport_Utility.ISteelProperty
    {
        // Rp0,2 МПа
        // Rm МПа
        // E, МПа
        // α 10^6
        // sigm

        double[] m_Rp;
        double[] m_Rm;
        double[] m_E;
        double[] m_Alpha;
        double[] m_Sigma;

        //double[] m_T;
        string m_Name;

        #region === Constructors ===

        public SteelProperty()
        {
            int len = 6;

            m_Rp = new double[len];
            m_Rm = new double[len];
            m_E = new double[len];
            m_Alpha = new double[len];
            m_Sigma = new double[len];

            //m_T = null;
            m_Name = "Noname";
        }

        public SteelProperty(double[] Rp, double[] Rm, double[] E, double[] Alpha, double[] Sigma, string Name)
        {
            int len = Rp.Length;

            m_Rp = new double[len];
            m_Rm = new double[len];
            m_E = new double[len];
            m_Alpha = new double[len];
            m_Sigma = new double[len];

            Rp.CopyTo(m_Rp, 0);
            Rm.CopyTo(m_Rm, 0);
            E.CopyTo(m_E, 0);
            Alpha.CopyTo(m_Alpha, 0);
            Sigma.CopyTo(m_Sigma, 0);

            //m_T = T;
            m_Name = Name;
        }

        #endregion

        #region === Public ===

        public double[] Rp
        {
            get { return m_Rp; }
            set
            {
                int len = value.Length;
                for (int i = 0; i < len; i++) m_Rp[i] = value[i];
            }
        }

        public double[] Rm
        {
            get { return m_Rm; }
            set
            {
                int len = value.Length;
                for (int i = 0; i < len; i++) m_Rm[i] = value[i];
            }
        }

        public double[] E
        {
            get { return m_E; }
            set
            {
                int len = value.Length;
                for (int i = 0; i < len; i++) m_E[i] = value[i];
            }
        }

        public double[] Alpha
        {
            get { return m_Alpha; }
            set
            {
                int len = value.Length;
                for (int i = 0; i < len; i++) m_Alpha[i] = value[i];
            }
        }

        public double[] Sigma
        {
            get { return m_Sigma; }
            set
            {
                int len = value.Length;
                for (int i = 0; i < len; i++) m_Sigma[i] = value[i];
            }
        }

        public string Name
        {
            get { return m_Name; }
            set { m_Name = value; }
        }

        #endregion
    }

    public class Set
    {
        List<int> m_set;

        public Set()
        {
            m_set = new List<int>();
        }

        public Set(int begin, int end)
        {
            m_set = new List<int>();
            AddRange(begin, end);
        }

        #region === public ===

        public void Add(int item)
        {
            m_set.Add(item);
        }

        public void AddRange(int begin, int end)
        {
            for (int x = begin; x <= end; x++)
            {
                m_set.Add(x);
            }
        }

        public bool In(int item)
        {
            return Array.IndexOf(m_set.ToArray(), item) >= 0;
        }

        #endregion
    }
}
