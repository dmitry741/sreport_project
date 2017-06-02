using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using PhisicsDependency;
using SReport_Utility;
using DbReader;

namespace StrengthReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region === members ===

        const int c_pto_count = 181;

        PhisicFeatureManager m_manager = new PhisicFeatureManager();
        System.Globalization.NumberFormatInfo m_nfi = null;
        System.Collections.Generic.List<string> m_pto_list = new System.Collections.Generic.List<string>();

        MyStringHashTable m_HashSteel = new MyStringHashTable();

        protected delegate string GDF(double val, int sign);
        protected GDF GetDoubleFormat = (x, sign) => Math.Round(x, sign).ToString();

        string m_lastError = string.Empty;
        double m_PressureCalc = 0;
        string m_source_path = string.Empty;

        // sets
        Set m_set_0408 = null;
        Set m_setBig = null;
        Set m_set212247 = null;

        // толщины плит
        double[] m_s = { 0.4, 0.5, 0.6, 0.7, 0.8, 0.9 };

        // hash tables
        MyDoubleHashTable m_hashTable43 = null;
        MyDoubleHashTable m_hashTable423 = null;
        MyDoubleHashTable m_hashTable422 = null;
        MyDoubleHashTable m_hashTablePlit = null;
        MyDoubleHashTable m_hashTable424 = null;
        MyDoubleHashTable m_hashTable51 = null;
        MyDoubleHashTable m_list52 = null;
        MyDoubleHashTable m_len52 = null;
        MyIntegerHashTable m_ptoHashLen = null;
        MyDoubleHashTable m_list53 = null;
        MyDoubleHashTable m_list54 = null;
        MyDoubleHashTable m_list55 = null;
        MyDoubleHashTable m_list56 = null;
        MyDoubleHashTable m_list57 = null;

        // material - plits
        SteelProperty m_PlitSt3 = null;
        SteelProperty m_Plit09G2C = null;

        // material - napr-up
        SteelProperty m_NaprUpSt2 = null;
        SteelProperty m_NaprUpSt3 = null;
        SteelProperty m_NaprUp20 = null;
        SteelProperty m_NaprUp20x13 = null;

        // material - napr-down
        SteelProperty m_NaprDownSt2 = null;
        SteelProperty m_NaprDownSt3 = null;
        SteelProperty m_NaprDown20 = null;
        SteelProperty m_NaprDown20x13 = null;

        // material - krepeg
        SteelProperty m_Krepeg40x = null;
        SteelProperty m_Krepeg35 = null;

        // material - rezba
        SteelProperty m_Rezba09G2C = null;
        SteelProperty m_Rezba20x13 = null;
        SteelProperty m_Rezba20 = null;

        // plates
        SteelProperty m_AISI = null;
        SteelProperty m_Titan = null;

        // steel 45
        SteelProperty m_Steel45 = null;

        // длины направляющих
        MyIntegerHashTable m_hashLen = new MyIntegerHashTable();

        // Main temperature
        double[] m_T = null;

        // messages
        System.Collections.ArrayList m_messages = new System.Collections.ArrayList();

        // флажок - true если ПТО > 47  и верхняя направляющая крепится встык, false в противном случае
        bool m_bNaprIn = true;

        // dictionary with steel
        System.Collections.Generic.Dictionary<string, SteelProperty> m_dicNaprUp = new System.Collections.Generic.Dictionary<string, SteelProperty>();
        System.Collections.Generic.Dictionary<string, SteelProperty> m_dicNaprDown = new System.Collections.Generic.Dictionary<string, SteelProperty>();
        System.Collections.Generic.Dictionary<string, SteelProperty> m_dicRezbaInNapr = new System.Collections.Generic.Dictionary<string, SteelProperty>();
        System.Collections.Generic.Dictionary<string, SteelProperty> m_dicGaika = new System.Collections.Generic.Dictionary<string, SteelProperty>();

        #endregion

        #region === public ===

        public string Source_Path
        {
            get
            {
                m_source_path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SReport";

                if (!Directory.Exists(m_source_path))
                    Directory.CreateDirectory(m_source_path);

                return m_source_path;
            }
        }

        #endregion

        #region === private ===

        private void StartSets()
        {
            m_set_0408 = new Set(0, 11);   
                     
            m_set212247 = new Set();
            m_set212247.AddRange(30, 47);
            m_set212247.AddRange(170, 175); // 26
            m_set212247.AddRange(154, 157); // 19
            m_set212247.AddRange(162, 169); // 19w
            m_set212247.AddRange(176, 178); // 31

            m_setBig = new Set();
            m_setBig.AddRange(48, 153); 
            m_setBig.AddRange(158, 161); // 53, 160
            m_setBig.AddRange(179, 180); // 229
        }

        private bool FillTables(ref PrintReport report)
        {
            bool result = false;

            do
            {
                if (!FillPersonal(ref report))
                    break;

                if (!FillTable1(ref report))
                    break;

                if (!FillTable2(ref report))
                    break;

                if (!FillTable3(ref report))
                    break;

                if (!FillTable4(ref report))
                    break;

                if (!FillTable5(ref report))
                    break;

                if (!FillTable6(ref report))
                    break;

                if (!FillTable8(ref report))
                    break;

                if (!FillTable9(ref report))
                    break;

                if (!FillTable10(ref report))
                    break;

                if (!FillTable11(ref report))
                    break;

                if (!FillTable12(ref report))
                    break;

                if (!FillTable13(ref report))
                    break;

                if (!FillTable14(ref report))
                    break;

                if (!FillTable15(ref report))
                    break;

                if (!FillTable16(ref report))
                    break;

                result = true;

            } while (false);

            return result;
        }

        private void FillDictSteel()
        {
            m_dicNaprUp.Add(KitConstant.Steel_st2, m_NaprUpSt2);
            m_dicNaprUp.Add(KitConstant.Steel_st3, m_NaprUpSt3);
            m_dicNaprUp.Add(KitConstant.Steel_20, m_NaprUp20);
            m_dicNaprUp.Add(KitConstant.Steel_20x13, m_NaprUp20x13);
            m_dicNaprUp.Add(KitConstant.Steel_45, m_Steel45);

            m_dicNaprDown.Add(KitConstant.Steel_st2, m_NaprDownSt2);
            m_dicNaprDown.Add(KitConstant.Steel_st3, m_NaprDownSt3);
            m_dicNaprDown.Add(KitConstant.Steel_20, m_NaprDown20);
            m_dicNaprDown.Add(KitConstant.Steel_20x13, m_NaprUp20x13);
            m_dicNaprDown.Add(KitConstant.Steel_45, m_Steel45);

            m_dicRezbaInNapr.Add(KitConstant.Steel_09G2C, m_Rezba09G2C);
            m_dicRezbaInNapr.Add(KitConstant.Steel_20x13, m_Rezba20x13);
            m_dicRezbaInNapr.Add(KitConstant.Steel_20, m_Rezba20);
            m_dicRezbaInNapr.Add(KitConstant.Steel_45, m_Steel45);

            m_dicGaika.Add(KitConstant.Steel_35, m_Krepeg35);
            m_dicGaika.Add(KitConstant.Steel_45, m_Steel45);
            m_dicGaika.Add(KitConstant.Steel_40x, m_Krepeg40x);
        }

        private SteelProperty GetNaprUp()
        {
            return m_dicNaprUp[comboBox7.Text];
        }

        private SteelProperty GetNaprDown()
        {
            return m_dicNaprDown[comboBox8.Text];
        }

        private SteelProperty RezbaInNapr()
        {
            return m_dicRezbaInNapr[comboBox12.Text];
        }

        private SteelProperty Gaika()
        {
            return m_dicGaika[comboBox10.Text];
        }

        private bool IsBig()
        {
            int ptoIndex = comboBox1.SelectedIndex;
            return m_setBig.In(ptoIndex);
        }

        private bool Is0408()
        {
            int ptoIndex = comboBox1.SelectedIndex;
            return m_set_0408.In(ptoIndex);
        }

        private bool Is212247()
        {
            int ptoIndex = comboBox1.SelectedIndex;
            return m_set212247.In(ptoIndex);
        }

        private void CreatePTOList()
        {
            // 4
            m_pto_list.Add("HH№04-O-10"); // 0
            m_pto_list.Add("HH№04-O-16"); // 1
            m_pto_list.Add("HH№04-O-25"); // 2

            m_pto_list.Add("HH№04-C-10"); // 3
            m_pto_list.Add("HH№04-C-16"); // 4
            m_pto_list.Add("HH№04-C-25"); // 5

            // 8
            m_pto_list.Add("HH№08-O-10"); // 6
            m_pto_list.Add("HH№08-O-16"); // 7
            m_pto_list.Add("HH№08-O-25"); // 8

            m_pto_list.Add("HH№08-C-10"); // 9
            m_pto_list.Add("HH№08-C-16"); // 10
            m_pto_list.Add("HH№08-C-25"); // 11

            // 7
            m_pto_list.Add("HH№07-O-10"); // 12
            m_pto_list.Add("HH№07-O-16"); // 13
            m_pto_list.Add("HH№07-O-25"); // 14

            m_pto_list.Add("HH№07-C-10"); // 15
            m_pto_list.Add("HH№07-C-16"); // 16
            m_pto_list.Add("HH№07-C-25"); // 17

            // 14
            m_pto_list.Add("HH№14-O-10"); // 18
            m_pto_list.Add("HH№14-O-16"); // 19
            m_pto_list.Add("HH№14-O-25"); // 20

            m_pto_list.Add("HH№14-C-10"); // 21
            m_pto_list.Add("HH№14-C-16"); // 22
            m_pto_list.Add("HH№14-C-25"); // 23

            // 20
            m_pto_list.Add("HH№20-O-10"); // 24
            m_pto_list.Add("HH№20-O-16"); // 25
            m_pto_list.Add("HH№20-O-25"); // 26

            m_pto_list.Add("HH№20-C-10"); // 27
            m_pto_list.Add("HH№20-C-16"); // 28
            m_pto_list.Add("HH№20-C-25"); // 29

            // 21
            m_pto_list.Add("HH№21-O-10"); // 30
            m_pto_list.Add("HH№21-O-16"); // 31
            m_pto_list.Add("HH№21-O-25"); // 32

            m_pto_list.Add("HH№21-C-10"); // 33
            m_pto_list.Add("HH№21-C-16"); // 34
            m_pto_list.Add("HH№21-C-25"); // 35

            // 22
            m_pto_list.Add("HH№22-O-10"); // 36
            m_pto_list.Add("HH№22-O-16"); // 37
            m_pto_list.Add("HH№22-O-25"); // 38

            m_pto_list.Add("HH№22-C-10"); // 39
            m_pto_list.Add("HH№22-C-16"); // 40
            m_pto_list.Add("HH№22-C-25"); // 41

            // 47
            m_pto_list.Add("HH№47-O-10"); // 42
            m_pto_list.Add("HH№47-O-16"); // 43
            m_pto_list.Add("HH№47-O-25"); // 44

            m_pto_list.Add("HH№47-C-10"); // 45
            m_pto_list.Add("HH№47-C-16"); // 46
            m_pto_list.Add("HH№47-C-25"); // 47

            // 41
            m_pto_list.Add("HH№41-O-10"); // 48
            m_pto_list.Add("HH№41-O-16"); // 49
            m_pto_list.Add("HH№41-O-25"); // 50

            m_pto_list.Add("HH№41-C-10");   // 51
            m_pto_list.Add("HH№41-C-16");   // 52
            m_pto_list.Add("HH№41-C-25");   // 53

            // 42
            m_pto_list.Add("HH№42-O-10");   // 54
            m_pto_list.Add("HH№42-O-16");   // 55
            m_pto_list.Add("HH№42-O-25");   // 56

            m_pto_list.Add("HH№42-C-10");   // 57
            m_pto_list.Add("HH№42-C-16");   // 58
            m_pto_list.Add("HH№42-C-25");   // 59

            // 62
            m_pto_list.Add("HH№62-O-10");   // 60
            m_pto_list.Add("HH№62-O-16");   // 61
            m_pto_list.Add("HH№62-O-25");   // 62

            m_pto_list.Add("HH№62-C-10");   // 63
            m_pto_list.Add("HH№62-C-16");   // 64
            m_pto_list.Add("HH№62-C-25");   // 65

            // 86
            m_pto_list.Add("HH№86-O-10");   // 66
            m_pto_list.Add("HH№86-O-16");   // 67
            m_pto_list.Add("HH№86-O-25");   // 68

            m_pto_list.Add("HH№86-C-10");   // 69
            m_pto_list.Add("HH№86-C-16");   // 70
            m_pto_list.Add("HH№86-C-25");   // 71

            // 110
            m_pto_list.Add("HH№110-O-10");   // 72
            m_pto_list.Add("HH№110-O-16");   // 73
            m_pto_list.Add("HH№110-O-25");   // 74

            m_pto_list.Add("HH№110-C-10");   // 75
            m_pto_list.Add("HH№110-C-16");   // 76
            m_pto_list.Add("HH№110-C-25");   // 77

            // 43
            m_pto_list.Add("HH№43-O-10");   // 78
            m_pto_list.Add("HH№43-O-16");   // 79
            m_pto_list.Add("HH№43-O-25");   // 80

            m_pto_list.Add("HH№43-C-10");   // 81
            m_pto_list.Add("HH№43-C-16");   // 82
            m_pto_list.Add("HH№43-C-25");   // 83

            // 65
            m_pto_list.Add("HH№65-O-10");   // 84
            m_pto_list.Add("HH№65-O-16");   // 85
            m_pto_list.Add("HH№65-O-25");   // 86

            m_pto_list.Add("HH№65-C-10");   // 87
            m_pto_list.Add("HH№65-C-16");   // 88
            m_pto_list.Add("HH№65-C-25");   // 89

            // 100
            m_pto_list.Add("HH№100-O-10");   // 90
            m_pto_list.Add("HH№100-O-16");   // 91
            m_pto_list.Add("HH№100-O-25");   // 92

            m_pto_list.Add("HH№100-C-10");   // 93
            m_pto_list.Add("HH№100-C-16");   // 94
            m_pto_list.Add("HH№100-C-25");   // 95

            // 130
            m_pto_list.Add("HH№130-O-10");   // 96
            m_pto_list.Add("HH№130-O-16");   // 97
            m_pto_list.Add("HH№130-O-25");   // 98

            m_pto_list.Add("HH№130-C-10");   // 99
            m_pto_list.Add("HH№130-C-16");   // 100
            m_pto_list.Add("HH№130-C-25");   // 101

            // 152
            m_pto_list.Add("HH№152-O-10");   // 102
            m_pto_list.Add("HH№152-O-16");   // 103
            m_pto_list.Add("HH№152-O-25");   // 104

            m_pto_list.Add("HH№152-C-10");   // 105
            m_pto_list.Add("HH№152-C-16");   // 106
            m_pto_list.Add("HH№152-C-25");   // 107

            // 220
            m_pto_list.Add("HH№220-O-10");   // 108
            m_pto_list.Add("HH№220-O-16");   // 109
            m_pto_list.Add("HH№220-O-25");   // 110

            m_pto_list.Add("HH№220-C-10");   // 111
            m_pto_list.Add("HH№220-C-16");   // 112
            m_pto_list.Add("HH№220-C-25");   // 113

            // 113
            m_pto_list.Add("HH№113-O-10");   // 114
            m_pto_list.Add("HH№113-O-16");   // 115
            m_pto_list.Add("HH№113-O-25");   // 116

            m_pto_list.Add("HH№113-C-10");   // 117
            m_pto_list.Add("HH№113-C-16");   // 118
            m_pto_list.Add("HH№113-C-25");   // 119

            // 81
            m_pto_list.Add("HH№81-O-10");   // 120
            m_pto_list.Add("HH№81-O-16");   // 121
            m_pto_list.Add("HH№81-O-25");   // 122

            m_pto_list.Add("HH№81-C-10");   // 123
            m_pto_list.Add("HH№81-C-16");   // 124
            m_pto_list.Add("HH№81-C-25");   // 125

            // 121
            m_pto_list.Add("HH№121-O-10");   // 126
            m_pto_list.Add("HH№121-O-16");   // 127
            m_pto_list.Add("HH№121-O-25");   // 128

            m_pto_list.Add("HH№121-C-10");   // 129
            m_pto_list.Add("HH№121-C-16");   // 130
            m_pto_list.Add("HH№121-C-25");   // 131

            // 188
            m_pto_list.Add("HH№188-O-10");   // 132
            m_pto_list.Add("HH№188-O-16");   // 133
            m_pto_list.Add("HH№188-O-25");   // 134

            m_pto_list.Add("HH№188-C-10");   // 135
            m_pto_list.Add("HH№188-C-16");   // 136
            m_pto_list.Add("HH№188-C-25");   // 137

            // 251
            m_pto_list.Add("HH№251-O-10");   // 138
            m_pto_list.Add("HH№251-O-16");   // 139
            m_pto_list.Add("HH№251-O-25");   // 140

            m_pto_list.Add("HH№251-C-10");   // 141
            m_pto_list.Add("HH№251-C-16");   // 142
            m_pto_list.Add("HH№251-C-25");   // 143

            // 145
            m_pto_list.Add("HH№145-O-10");   // 144
            m_pto_list.Add("HH№145-O-16");   // 145

            m_pto_list.Add("HH№145-C-10");   // 146
            m_pto_list.Add("HH№145-C-16");   // 147

            // 210
            m_pto_list.Add("HH№210-O-10");   // 148
            m_pto_list.Add("HH№210-O-16");   // 149

            m_pto_list.Add("HH№210-C-10");   // 150
            m_pto_list.Add("HH№210-C-16");   // 151

            // 201
            m_pto_list.Add("HH№201-O-10");   // 152
            m_pto_list.Add("HH№201-C-10");   // 153

            // 19
            m_pto_list.Add("HH№19-O-10");   // 154
            m_pto_list.Add("HH№19-O-16");   // 155

            m_pto_list.Add("HH№19-C-10");   // 156
            m_pto_list.Add("HH№19-C-16");   // 157

            // 53
            m_pto_list.Add("HH№53-O-10");   // 158
            m_pto_list.Add("HH№53-C-10");   // 159

            // 160
            m_pto_list.Add("HH№160-O-10");   // 160
            m_pto_list.Add("HH№160-C-10");   // 161

            // 19w
            m_pto_list.Add("HH№19-O-25"); // 162
            m_pto_list.Add("HH№19-C-25"); // 163

            m_pto_list.Add("HH№19w-O-10"); // 164
            m_pto_list.Add("HH№19w-O-16"); // 165
            m_pto_list.Add("HH№19w-C-10"); // 166
            m_pto_list.Add("HH№19w-C-16"); // 167
            m_pto_list.Add("HH№19w-O-25"); // 168
            m_pto_list.Add("HH№19w-C-25"); // 169

            // 26
            m_pto_list.Add("HH№26-O-10"); // 170
            m_pto_list.Add("HH№26-O-16"); // 171
            m_pto_list.Add("HH№26-O-25"); // 172
            m_pto_list.Add("HH№26-C-10"); // 173
            m_pto_list.Add("HH№26-C-16"); // 174
            m_pto_list.Add("HH№26-C-25"); // 175

            // XG-31
            m_pto_list.Add("XG31-C-10"); // 176
            m_pto_list.Add("XG31-C-16"); // 177
            m_pto_list.Add("XG31-C-25"); // 178

            // 229
            m_pto_list.Add("HH№229-O-10"); // 179
            m_pto_list.Add("HH№229-C-10"); // 180
        }

        private string GetPrePlaneNumber()
        {
            string sPto = comboBox1.Text;
            int pos1 = sPto.IndexOf('№');
            int pos2 = sPto.IndexOf('-');
            int length = sPto.Length;
            string s = sPto.Substring(pos1 + 1, pos2 - pos1 - 1);

            while (s.Length < 3)
            {
                s = "0" + s;
            }

            return s;
        }

        private string GetPtoOriginalName()
        {
            string sOriginalName = comboBox1.Text;
            int pos = sOriginalName.IndexOf('-');

            if (pos > 0)
            {
                sOriginalName = sOriginalName.Remove(pos);
            }

            return sOriginalName;
        }

        private string GetPtoString()
        {
            string sPto = comboBox1.Text;
            int pos = sPto.IndexOf('-');

            if (pos > 0)
            {
                sPto = sPto.Remove(pos + 1);
                sPto += textBox4.Text;
                sPto += "PP";
            }

            return sPto;
        }

        private bool FillTable1(ref SReport_Utility.PrintReport report)
        {
            // phrase1
            if (comboBox13.SelectedIndex != comboBox14.SelectedIndex)
            {
                string s = "1.2 Данный теплообменник рассчитан для применения при двух различных режимах. Для проведения расчета на прочность выбран наихудший режим.";
                report.SetValue("Table1_Phrase1", s);
            }

            report.SetValue("PTO", GetPtoString());
            report.SetValue("PTO_OriginalName", GetPtoOriginalName());
            report.SetValue("PTO_Plane_Number", textBox4.Text);
            report.SetValue("PTO_Full_Plane_Number", GetPrePlaneNumber() + "-" + textBox4.Text);

            report.SetValue("Medium_Hot_Name", txtMediumHot.Text);
            report.SetValue("Medium_Cold_Name", txtMediumCold.Text);
            report.SetValue("Medium_Hot_MaxT", textBox7.Text);
            report.SetValue("Medium_Cold_MaxT", textBox12.Text);

            double MinT;
            double.TryParse(textBox8.Text, out MinT);

            report.SetValue("Medium_MinT", (MinT < 0) ? "минус " + Math.Abs(MinT).ToString() : MinT.ToString());
            report.SetValue("Medium_WallT", textBox11.Text);

            report.SetValue("Weight_Netto", textBox10.Text);
            report.SetValue("Weight_Brutto", textBox9.Text);

            report.SetValue("Pmax", textBox1.Text);
            report.SetValue("Table1_Period", "20");

            // === plate count ===
            int nPlates;
            int.TryParse(textBox3.Text, out nPlates);

            report.SetValue("PTO_Plates", nPlates.ToString());

            nPlates--;
            report.SetValue("PTO_Plates_1", nPlates.ToString());
            // ===================

            Bitmap bitmap1 = SimpleImageProccessing.Zoom(Properties.Resources.PTO, 241, 350);
            report.SetPicture("Picture1", bitmap1);

            Bitmap bitmap2 = SimpleImageProccessing.Zoom(Properties.Resources.PTO2, 1021, 1165);
            report.SetPicture("Picture2", bitmap2);

            return true;
        }

        private bool FillPersonal(ref SReport_Utility.PrintReport report)
        {
            report.SetValue("People_Developer", textBox13.Text);
            report.SetValue("People_Auditor", textBox14.Text);
            report.SetValue("People_NachPodrz", textBox16.Text);
            report.SetValue("People_Ncontr", textBox15.Text);
            report.SetValue("People_Utverdil", textBox17.Text);

            return true;
        }

        private string GetSteelValue(string Steel)
        {
            string value = Steel;
            value += " ";

            string val2 = m_HashSteel.GetValue2(Steel);

            if (val2.Length > 0)
            {
                value += val2;
                value += " ";
            }

            value += Convert.ToChar(10); // new line
            value += m_HashSteel.GetValue1(Steel);

            return value;
        }

        private void Fill_PredelTech(ref SReport_Utility.PrintReport report, double T1, double T2)
        {                        
            double[] X = m_T;
            double[] Y = null;
            bool Formated = false;
            int sign = 1;

            // 1. плита неподвижная
            SteelProperty steel = (KitConstant.Steel_st3 == comboBox5.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.Rp;
            PhisicsDependency.Lines lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_11", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_12", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 2. Направляющая верхняя
            steel = GetNaprUp();

            Y = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_21", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_22", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 3. Направляющая нижняя
            steel = GetNaprDown();

            Y = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_31", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_32", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 4. резьбовая часть в напрявляющих
            steel = RezbaInNapr();

            Y = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_41", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_42", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 5. шпилька стяжная
            steel = m_Krepeg40x;
            Y = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_51", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_52", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 6. болт под фланцы
            steel = m_Krepeg40x;
            Y = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_61", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_62", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 7. гайка
            steel = Gaika();
            Y = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_71", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_72", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 8. плита прижимная
            steel = (KitConstant.Steel_st3 == comboBox6.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelTech_81", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelTech_82", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));
        }

        private void Fill_PredelProch(ref SReport_Utility.PrintReport report, double T1, double T2)
        {            
            double[] X = m_T;
            double[] Y = null;
            int sign = 1;
            bool Formated = false;

            // 1. плита неподвижная, плита прижимная
            SteelProperty steel = (KitConstant.Steel_st3 == comboBox5.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.Rm;
            PhisicsDependency.Lines lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_11", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_12", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 2. Направляющая верхняя
            steel = GetNaprUp();

            Y = steel.Rm;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_21", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_22", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 3. Направляющая нижняя
            steel = GetNaprDown();

            Y = steel.Rm;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_31", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_32", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 4. резьбовая часть в напрявляющих
            steel = RezbaInNapr();

            Y = steel.Rm;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_41", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_42", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 5. шпилька стяжная
            steel = m_Krepeg40x;
            Y = steel.Rm;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_51", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_52", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 6. болт под фланцы
            steel = m_Krepeg40x;
            Y = steel.Rm;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_61", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_62", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 7. гайка
            steel = Gaika();
            Y = steel.Rm;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_71", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_72", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));

            // 8. плита прижимная
            steel = (KitConstant.Steel_st3 == comboBox6.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.Rm;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_PredelProch_81", DigitalProcess.MathTrancate(lines.Value(T1), sign, Formated));
            report.SetValue("Table2_PredelProch_82", DigitalProcess.MathTrancate(lines.Value(T2), sign, Formated));
        }

        private void Fill_Sigma(ref SReport_Utility.PrintReport report, double T1, double T2)
        {
            double[] X = m_T;
            double[] Y1 = null;
            double[] Y2 = null;
            double Np = 1.5;
            double Nm = 2.4;
            double val1, val2;
            int sign = 1;
            bool Formated = false;

            // 1. плита неподвижная
            SteelProperty steel = (KitConstant.Steel_st3 == comboBox5.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y1 = steel.Rp;
            Y2 = steel.Rm;

            PhisicsDependency.Lines lines1 = new PhisicsDependency.Lines(X, Y1);
            PhisicsDependency.Lines lines2 = new PhisicsDependency.Lines(X, Y2);

            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            report.SetValue("Table2_Sigma_11", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_12", DigitalProcess.MathTrancate(val2, sign, Formated));

            // 2. Направляющая верхняя
            steel = GetNaprUp();

            Y1 = steel.Rp;
            Y2 = steel.Rm;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            lines2 = new PhisicsDependency.Lines(X, Y2);
            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            report.SetValue("Table2_Sigma_21", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_22", GetDoubleFormat(val2, 1));

            // 3. Направляющая нижняя
            steel = GetNaprDown();

            Y1 = steel.Rp;
            Y2 = steel.Rm;

            lines1 = new PhisicsDependency.Lines(X, Y1);
            lines2 = new PhisicsDependency.Lines(X, Y2);

            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            report.SetValue("Table2_Sigma_31", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_32", GetDoubleFormat(val2, 1));

            // 4. резьбовая часть в напрявляющих (крепеж)
            steel = RezbaInNapr();

            Y1 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            report.SetValue("Table2_Sigma_41", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_42", DigitalProcess.MathTrancate(val2, sign, Formated));

            // 5. шпилька стяжная (крепеж)
            steel = m_Krepeg40x;
            Y1 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            report.SetValue("Table2_Sigma_51", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_52", DigitalProcess.MathTrancate(val2, sign, Formated));

            // 6. болт под фланцы (крепеж)
            steel = m_Krepeg40x;
            Y1 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            report.SetValue("Table2_Sigma_61", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_62", DigitalProcess.MathTrancate(val2, sign, Formated));

            // 7. гайка (крепеж)
            steel = Gaika();
            Y1 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            report.SetValue("Table2_Sigma_71", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_72", DigitalProcess.MathTrancate(val2, sign, Formated));

            // 8. плита прижимная
            steel = (KitConstant.Steel_st3 == comboBox6.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y1 = steel.Rp;
            Y2 = steel.Rm;

            lines1 = new PhisicsDependency.Lines(X, Y1);
            lines2 = new PhisicsDependency.Lines(X, Y2);

            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            report.SetValue("Table2_Sigma_81", DigitalProcess.MathTrancate(val1, sign, Formated));
            report.SetValue("Table2_Sigma_82", DigitalProcess.MathTrancate(val2, sign, Formated));
        }

        private void Fill_ModuleUpr(ref SReport_Utility.PrintReport report, double T1, double T2)
        {            
            double[] X = m_T;
            double[] Y = null;
            double d = 100000;
            int sign = 2;
            string S1, S2;
            double val1, val2;

            // 1. плита неподвижная
            SteelProperty steel = (KitConstant.Steel_st3 == comboBox5.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.E;
            PhisicsDependency.Lines lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_11", S1);
            report.SetValue("Table2_ModuleUpr_12", S2);

            // 2. Направляющая верхняя
            steel = GetNaprUp();

            Y = steel.E;
            lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_21", S1);
            report.SetValue("Table2_ModuleUpr_22", S2);

            // 3. Направляющая нижняя
            steel = GetNaprDown();

            Y = steel.E;
            lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_31", S1);
            report.SetValue("Table2_ModuleUpr_32", S2);

            // 4. резьбовая часть в направляющих
            steel = RezbaInNapr();

            Y = steel.E;
            lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_41", S1);
            report.SetValue("Table2_ModuleUpr_42", S2);

            // 5. шпилька стяжная
            steel = m_Krepeg40x;
            Y = steel.E;
            lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_51", S1);
            report.SetValue("Table2_ModuleUpr_52", S2);

            // 6. болт под фланцы
            steel = m_Krepeg40x;
            Y = steel.E;
            lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_61", S1);
            report.SetValue("Table2_ModuleUpr_62", S2);

            // 7. гайка
            steel = Gaika();
            Y = steel.E;
            lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_71", S1);
            report.SetValue("Table2_ModuleUpr_72", S2);

            // 8. плита прижимная
            steel = (KitConstant.Steel_st3 == comboBox6.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.E;
            lines = new PhisicsDependency.Lines(X, Y);
            val1 = lines.Value(T1) / d;
            val2 = lines.Value(T2) / d;
            S1 = DigitalProcess.MathTrancate(val1, sign, true);
            S2 = DigitalProcess.MathTrancate(val2, sign, true);

            report.SetValue("Table2_ModuleUpr_81", S1);
            report.SetValue("Table2_ModuleUpr_82", S2);
        }

        private void Fill_LinearExt(ref SReport_Utility.PrintReport report, double T1, double T2)
        {            
            double[] X = m_T;
            double[] Y = null;
            int sign = 2;

            // 1. плита неподвижная
            SteelProperty steel = (KitConstant.Steel_st3 == comboBox5.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.Alpha;
            PhisicsDependency.Lines lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_12", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));

            // 2. Направляющая верхняя
            steel = GetNaprUp();

            Y = steel.Alpha;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_22", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));

            // 3. Направляющая нижняя
            steel = GetNaprDown();

            Y = steel.Alpha;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_32", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));

            // 4. резьбовая часть в напрявляющих
            steel = RezbaInNapr();

            Y = steel.Alpha;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_42", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));

            // 5. шпилька стяжная
            steel = m_Krepeg40x;
            Y = steel.Alpha;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_52", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));

            // 6. болт под фланцы
            steel = m_Krepeg40x;
            Y = steel.Alpha;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_62", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));

            // 7. гайка
            steel = Gaika();

            Y = steel.Alpha;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_72", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));

            // 8. плита прижимная
            steel = (KitConstant.Steel_st3 == comboBox6.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y = steel.Alpha;
            lines = new PhisicsDependency.Lines(X, Y);

            report.SetValue("Table2_LinearExt_82", DigitalProcess.MathTrancate(lines.Value(T2), sign, true));
        }

        private bool FillTable2(ref SReport_Utility.PrintReport report)
        {
            // 1. плита неподвижная
            report.SetValue("Table2_Material_Row1", GetSteelValue(comboBox5.Text));

            // 2. напрявляющая верхняя
            report.SetValue("Table2_Material_Row2", GetSteelValue(comboBox7.Text));

            // 3. направляющая нижняя
            report.SetValue("Table2_Material_Row3", GetSteelValue(comboBox8.Text));

            // 4. резьбовая часть в напрявляющих
            report.SetValue("Table2_Material_Row4", GetSteelValue(comboBox12.Text));

            // 5. шпилька стяжная
            report.SetValue("Table2_Material_Row5", GetSteelValue(comboBox9.Text));

            // 6. болт под фланцы
            report.SetValue("Table2_Material_Row6", GetSteelValue(comboBox11.Text));

            // 7. гайка
            report.SetValue("Table2_Material_Row7", GetSteelValue(comboBox10.Text));

            // 8. плита прижимная
            report.SetValue("Table2_Material_Row8", GetSteelValue(comboBox6.Text));

            // R18
            string r18 = (comboBox16.SelectedIndex < 3) ? "Болт 13, 14, 15, 16; болты крепления фланцев подвода-отвода рабочих сред." : "Болт 13, 14, 15, 16;";
            report.SetValue("Table2_R18", r18);

            // расчетная температура
            double T1 = KitConstant.HydroTest_T1;
            double T2 = 0;

            double.TryParse(textBox11.Text, out T2);

            report.SetValue("Table2_T1", T1.ToString());
            report.SetValue("Table2_T2", T2.ToString());

            // предел текучести
            Fill_PredelTech(ref report, T1, T2);

            // предел прочности
            Fill_PredelProch(ref report, T1, T2);

            // Номинальное допускаемое напряжение (сигма)
            Fill_Sigma(ref report, T1, T2);

            // Модуль упругости Eт 10-5, МПа
            Fill_ModuleUpr(ref report, T1, T2);

            // Коэффициент линейного расширения
            Fill_LinearExt(ref report, T1, T2);
            
            return true;
        }

        private bool FillTable3(ref SReport_Utility.PrintReport report)
        {
            double[] X = m_T;
            double[] Y1 = null;
            double[] Y2 = null;
            double Np = 1.5;
            double Nm = 2.4;
            double val1, val2;
            double Rmin = double.MaxValue;
            double Rc = 0;
            double SigmaTH = 1;
            double SigmaT = 1;            

            // расчетная температура
            double T1 = KitConstant.HydroTest_T1;
            double T2 = 0;

            double.TryParse(textBox11.Text, out T2);

            // 1. плита неподвижная
            SteelProperty steel = (KitConstant.Steel_st3 == comboBox5.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y1 = steel.Rp;
            Y2 = steel.Rm;
            PhisicsDependency.Lines lines1 = new PhisicsDependency.Lines(X, Y1);
            PhisicsDependency.Lines lines2 = new PhisicsDependency.Lines(X, Y2);
            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            Rc = val1 / val2;

            Rmin = Rc;
            SigmaTH = val1;
            SigmaT = val2;

            // 2. Направляющая верхняя
            steel = GetNaprUp();

            Y1 = steel.Rp;
            Y2 = steel.Rm;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            lines2 = new PhisicsDependency.Lines(X, Y2);
            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            Rc = val1 / val2;

            if (Rc < Rmin)
            {
                Rmin = Rc;
                SigmaTH = val1;
                SigmaT = val2;
            }

            // 3. Направляющая нижняя
            steel = GetNaprDown();

            Y1 = steel.Rp;
            Y2 = steel.Rm;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            lines2 = new PhisicsDependency.Lines(X, Y2);
            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            Rc = val1 / val2;

            if (Rc < Rmin)
            {
                Rmin = Rc;
                SigmaTH = val1;
                SigmaT = val2;
            }

            // 4. резьбовая часть в напрявляющих (крепеж)
            steel = RezbaInNapr();

            Y1 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            Rc = val1 / val2;

            if (Rc < Rmin)
            {
                Rmin = Rc;
                SigmaTH = val1;
                SigmaT = val2;
            }

            // 5. шпилька стяжная (крепеж)
            steel = m_Krepeg40x;
            Y1 = steel.Rp;
            Y2 = steel.Rm;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            Rc = val1 / val2;

            if (Rc < Rmin)
            {
                Rmin = Rc;
                SigmaTH = val1;
                SigmaT = val2;
            }

            // 6. болт под фланцы (крепеж)
            steel = m_Krepeg40x;
            Y1 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            Rc = val1 / val2;

            if (Rc < Rmin)
            {
                Rmin = Rc;
                SigmaTH = val1;
                SigmaT = val2;
            }

            // 7. гайка (крепеж)
            steel = Gaika();
            Y1 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            val1 = lines1.Value(T1) / 2;
            val2 = lines1.Value(T2) / 2;

            Rc = val1 / val2;

            if (Rc < Rmin)
            {
                Rmin = Rc;
                SigmaTH = val1;
                SigmaT = val2;
            }

            // 8. плита прижимная
            steel = (KitConstant.Steel_st3 == comboBox6.Text) ? m_PlitSt3 : m_Plit09G2C;

            Y1 = steel.Rp;
            Y2 = steel.Rm;
            lines1 = new PhisicsDependency.Lines(X, Y1);
            lines2 = new PhisicsDependency.Lines(X, Y2);
            val1 = Math.Min(lines1.Value(T1) / Np, lines2.Value(T1) / Nm);
            val2 = Math.Min(lines1.Value(T2) / Np, lines2.Value(T2) / Nm);

            Rc = val1 / val2;

            if (Rc < Rmin)
            {
                Rmin = Rc;
                SigmaTH = val1;
                SigmaT = val2;
            }

            report.SetValue("Table3_SigmaTH", DigitalProcess.MathTrancate(SigmaTH, 2, false));
            report.SetValue("Table3_SigmaT", DigitalProcess.MathTrancate(SigmaT, 2, false));

            // расчетное давление
            double P;
            double.TryParse(textBox1.Text, out P);
            report.SetValue("Table3_P", P.ToString());

            double Ph = 1.25 * P * Rmin;
            report.SetValue("Table3_PH", DigitalProcess.MathTrancate(Ph, 2, false));

            // давление гидроиспытаний
            double.TryParse(textBox2.Text, out P);
            report.SetValue("Presuure_HydroTest", DigitalProcess.MathTrancate(P, 1, true));

            //Pressure_Calc
            m_PressureCalc = P + 0.1;
            report.SetValue("Pressure_Calc", DigitalProcess.MathTrancate(m_PressureCalc, 1, true));

            return true;
        }

        private bool FillTable4(ref SReport_Utility.PrintReport report)
        {
            MyDoubleHashTable hashTable = m_hashTablePlit;
            string ptoName = comboBox1.Text;

            // "Bm", "B", "Am", "A", "b", "Bsh", "d", "z", "c111", "c112", "c12", "c2", "S1", "S2";

            double B = hashTable.GetValue(ptoName, "B");
            report.SetValue("Table4_B", B.ToString());

            double bb = hashTable.GetValue(ptoName, "b1");
            report.SetValue("Table_bb", bb.ToString());

            double Bsh = hashTable.GetValue(ptoName, "Bsh");
            report.SetValue("Table_Bsh", Bsh.ToString());

            double A = hashTable.GetValue(ptoName, "A");
            report.SetValue("Table_A", A.ToString());

            // Formula Y
            double Y = 1.41 / Math.Sqrt(1 + (B * B) / (A * A));
            report.SetValue("Table4_Y", DigitalProcess.MathTrancate(Y, 2, true));

            // Formula Km
            double Km = Bsh / (2 * (B + bb));
            report.SetValue("Table4_Km", DigitalProcess.MathTrancate(Km, 2, true));

            double d = hashTable.GetValue(ptoName, "d");
            report.SetValue("Table_d", d.ToString());

            double z = hashTable.GetValue(ptoName, "z");
            report.SetValue("Table_z", z.ToString());

            // Formula Ko
            double val = z * d / Bsh;
            double KoStat = Math.Sqrt(1 + val + val * val);
            double KoPri = 0.98; // 1
            report.SetValue("Table4_Ko", DigitalProcess.MathTrancate(KoStat, 3, true));

            double c111 = hashTable.GetValue(ptoName, "c111");
            report.SetValue("Table_c111", c111.ToString());

            double c112 = hashTable.GetValue(ptoName, "c112");
            report.SetValue("Table_c112", c112.ToString());

            double c12 = hashTable.GetValue(ptoName, "c12");
            report.SetValue("Table_c12", c12.ToString());

            double c2 = hashTable.GetValue(ptoName, "c2");
            report.SetValue("Table_c2", c12.ToString());

            // Formula C
            double c_stat = c111 + c12 + c2;
            report.SetValue("Table4_c_stat", c_stat.ToString());

            double c_pri = c112 + c12 + c2;
            report.SetValue("Table4_c_pri", c_pri.ToString());

            double S1 = hashTable.GetValue(ptoName, "S1");
            report.SetValue("Table_S1", S1.ToString());

            double S2 = hashTable.GetValue(ptoName, "S2");
            report.SetValue("Table_S2", S2.ToString());

            // Sigma
            double Sigma_stat = 0;
            double.TryParse(report.GetValue("Table2_Sigma_12"), out Sigma_stat);

            double Sigma_pri = 0;
            double.TryParse(report.GetValue("Table2_Sigma_82"), out Sigma_pri);

            report.SetValue("Table4_Sigma", Sigma_stat.ToString());
            report.SetValue("Table4_Sigma_pri", Sigma_pri.ToString());            

            // расчетная температура
            double T1 = KitConstant.HydroTest_T1;
            double[] X = m_T;
            double[] Y1 = null;

            bool Formated = false;
            int sign = 1;

            // 1. плита неподвижная
            SteelProperty steel = (KitConstant.Steel_st3 == comboBox5.Text) ? m_PlitSt3 : m_Plit09G2C;
            Y1 = steel.Rp;
            PhisicsDependency.Lines lines = new PhisicsDependency.Lines(X, Y1);

            double SigmaH_stat = lines.Value(T1) / 1.1;
            report.SetValue("Table4_SigmaH", DigitalProcess.MathTrancate(SigmaH_stat, sign, Formated));

            // 2. плита прижимная
            steel = (KitConstant.Steel_st3 == comboBox6.Text) ? m_PlitSt3 : m_Plit09G2C;
            Y1 = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y1);

            double SigmaH_pri = lines.Value(T1) / 1.1;
            report.SetValue("Table4_SigmaH_pri", DigitalProcess.MathTrancate(SigmaH_pri, sign, Formated));

            double fi = 1;
            double P;
            double.TryParse(textBox1.Text, out P);
            double Ph = m_PressureCalc;

            // Sr1
            double Sr1_stat = KoStat * Km * Y * B * Math.Sqrt(P / (fi * Sigma_stat));
            double Sr1_pri = KoPri * Km * Y * B * Math.Sqrt(P / (fi * Sigma_pri));

            // Sr2
            double Sr2_stat = KoStat * Km * Y * B * Math.Sqrt(Ph / (fi * SigmaH_stat));
            double Sr2_pri = KoPri * Km * Y * B * Math.Sqrt(Ph / (fi * SigmaH_pri));

            Formated = true;
            sign = 2;

            report.SetValue("Table4_SR1", DigitalProcess.MathTrancate(Sr1_stat, sign, Formated));
            report.SetValue("Table4_SR2", DigitalProcess.MathTrancate(Sr2_stat, sign, Formated));
            report.SetValue("Table4_SR12", DigitalProcess.MathTrancate(Sr1_pri, sign, Formated));
            report.SetValue("Table4_SR22", DigitalProcess.MathTrancate(Sr2_pri, sign, Formated));

            report.SetValue("Table4_SR11_plus_C", DigitalProcess.MathTrancate(Sr1_stat + c_stat, sign, Formated));
            report.SetValue("Table4_SR21_plus_C", DigitalProcess.MathTrancate(Sr2_stat + c_stat, sign, Formated));
            report.SetValue("Table4_SR12_plus_C", DigitalProcess.MathTrancate(Sr1_pri + c_pri, sign, Formated));
            report.SetValue("Table4_SR22_plus_C", DigitalProcess.MathTrancate(Sr2_pri + c_pri, sign, Formated));

            // условие прочности
            bool Condition = (S1 > Sr1_stat + c_stat);
            Condition = Condition && (S1 > Sr2_stat + c_stat);
            Condition = Condition && (S2 > Sr1_pri + c_pri);
            Condition = Condition && (S2 > Sr2_pri + c_pri);

            if (!Condition)
            {
                m_messages.Add("Таблица 4, стр. 10");
            }

            return true;
        }

        private bool FillTable5(ref SReport_Utility.PrintReport report)
        {
            MyDoubleHashTable hashTable422 = m_hashTable422;
            MyDoubleHashTable hashTable43 = m_hashTable43;
            string ptoName = comboBox1.Text;

            // Epsilon
            double Epsilon = hashTable43.GetValue(ptoName, "epsi") * 100;
            report.SetValue("Table5_Epsilon", GetDoubleFormat(Epsilon, 0));

            // A
            double A = hashTable43.GetValue(ptoName, "Am");
            report.SetValue("Table5_Am", GetDoubleFormat(A, 0));

            // B
            double B = hashTable43.GetValue(ptoName, "Bm");
            report.SetValue("Table5_Bm", GetDoubleFormat(B, 0));

            // b
            double b = hashTable43.GetValue(ptoName, "b");
            report.SetValue("Table5_b", DigitalProcess.MathTrancate(b, 1, false));

            // q0
            int MediumHotIndex = comboBox13.SelectedIndex;
            int MediumColdIndex = comboBox14.SelectedIndex;
            double q0_Hot, q0_Cold;
            string sQuery;

            if (MediumHotIndex == 0) // liq
            {
                sQuery = "q0liq";
            }
            else if (MediumHotIndex == 1) // vap
            {
                sQuery = "q0vap";
            }
            else // air
            {
                sQuery = "q0air";
            }

            q0_Hot = hashTable422.GetValue(ptoName, sQuery);

            if (MediumColdIndex == 0) // liq
            {
                sQuery = "q0liq";
            }
            else if (MediumColdIndex == 1) // vap
            {
                sQuery = "q0vap";
            }
            else // air
            {
                sQuery = "q0air";
            }

            q0_Cold = hashTable422.GetValue(ptoName, sQuery);

            double q0 = Math.Max(q0_Hot, q0_Cold);
            report.SetValue("Table5_Q0", GetDoubleFormat(q0, 0));

            // Fob
            double Fob = 2 * (A + B) * b * q0;
            report.SetValue("Table5_Fob", DigitalProcess.MathTrancate(Convert.ToInt32(Fob), 2).ToString());

            // P
            double P;
            double.TryParse(textBox1.Text, out P);
            report.SetValue("Table5_P", DigitalProcess.MathTrancate(P, 1, true));

            // Ph
            report.SetValue("Table5_PH", DigitalProcess.MathTrancate(m_PressureCalc, 1, true));

            // Eta
            double Eta = hashTable422.GetValue(ptoName, "eta");
            report.SetValue("Table5_Eta", DigitalProcess.MathTrancate(Eta, 1, true));

            // Fp
            double Fp = A * B * P;
            report.SetValue("Table5_Fp", DigitalProcess.MathTrancate(Convert.ToInt32(Fp), 2).ToString());

            // Fd
            double Fd = Fob + 0.1 * Fp;
            report.SetValue("Table5_Fd", DigitalProcess.MathTrancate(Convert.ToInt32(Fd), 2).ToString());

            // Hi
            double Hi;
            double[] X1 = { 20, 200 };
            double[] Y1 = { 1.0, 1.5 };            

            PhisicsDependency.Lines lines1 = new PhisicsDependency.Lines(X1, Y1);

            double T2;
            double.TryParse(textBox11.Text, out T2);

            Hi = lines1.Value(T2);
            report.SetValue("Table5_Hi", DigitalProcess.MathTrancate(Hi, 2, false));

            // m
            double m_Hot, m_Cold;

            if (MediumHotIndex == 0) // liq
            {
                sQuery = "mliq";
            }
            else if (MediumHotIndex == 1) // vap
            {
                sQuery = "mvap";
            }
            else // air
            {
                sQuery = "mair";
            }

            m_Hot = hashTable422.GetValue(ptoName, sQuery);

            if (MediumColdIndex == 0) // liq
            {
                sQuery = "mliq";
            }
            else if (MediumColdIndex == 1) // vap
            {
                sQuery = "mvap";
            }
            else // air
            {
                sQuery = "mair";
            }

            m_Cold = hashTable422.GetValue(ptoName, sQuery);

            double m = Math.Max(m_Hot, m_Cold);
            report.SetValue("Table5_m", DigitalProcess.MathTrancate(m, 1, true));

            // F2
            double F2 = 2 * (A + B) * b * Hi * m * P;
            report.SetValue("Table5_F2", DigitalProcess.MathTrancate(Convert.ToInt32(F2), 2).ToString());

            // F2 plus
            double F2_plus = F2 + (1 - Eta) * Fp;
            report.SetValue("Table5_F2_plus", DigitalProcess.MathTrancate(Convert.ToInt32(F2_plus), 2).ToString());

            // Fph
            double Fph = A * B * m_PressureCalc;
            report.SetValue("Table5_Fph", DigitalProcess.MathTrancate(Convert.ToInt32(Fph), 2).ToString());

            // F2h
            double F2h = 0.8 * 2 * (A + B) * b * Hi * m * m_PressureCalc;
            report.SetValue("Table5_F2h", DigitalProcess.MathTrancate(Convert.ToInt32(F2h), 2).ToString());

            // F2h_plus
            double F2h_plus = F2h + (1 - Eta) * Fph;
            report.SetValue("Table5_F2h_plus", DigitalProcess.MathTrancate(Convert.ToInt32(F2h_plus), 2).ToString());

            // z
            double z = hashTable422.GetValue(ptoName, "z");
            report.SetValue("Table5_z", GetDoubleFormat(z, 0));

            // d0
            double d0 = hashTable422.GetValue(ptoName, "d0");
            report.SetValue("Table5_d0", GetDoubleFormat(d0, 0));

            // ksi
            double ksi = hashTable422.GetValue(ptoName, "ksi");
            report.SetValue("Table5_ksi", DigitalProcess.MathTrancate(ksi, 2, false));

            // Mk
            double max = Fd;

            if (F2_plus > max) max = F2_plus;
            if (F2h_plus > max) max = F2h_plus;

            max *= 1.1;

            double Mk = max * ksi * d0 / z;
            int nMk = DigitalProcess.MathTrancateInc(Convert.ToInt32(Mk), 2);
            report.SetValue("Table5_MomentK", nMk.ToString());

            // Fow
            double Fow = nMk * z / (ksi * d0);
            report.SetValue("Table5_Fow", DigitalProcess.MathTrancate(Convert.ToInt32(Fow), 2).ToString());

            // === correct 7 april 2010 ===

            max = Fd;

            if (F2_plus > max) max = F2_plus;
            if (F2h_plus > max) max = F2h_plus;
            if (Fow > max) max = Fow;

            // ============================

            // Fw
            //double Fw = Fow;
            double Fw = max;
            report.SetValue("Table5_Fw", DigitalProcess.MathTrancate(Convert.ToInt32(Fw), 2).ToString());

            // dc
            int dc = 0;
            report.SetValue("Table5_dc", dc.ToString());

            // SigmaW ( шпилька стяжная (крепеж) )
            SteelProperty steel = m_Krepeg40x;
            double[] X2 = m_T;
            double[] Y2 = steel.Rp;
            lines1 = new PhisicsDependency.Lines(X2, Y2);
            double SigmaW = lines1.Value(T2) / 2;
            report.SetValue("Table5_SigmaW", DigitalProcess.MathTrancate(SigmaW, 2, false));

            // dw
            double dw = Math.Sqrt(1.27 * Fw / (z * SigmaW) + dc * dc);
            report.SetValue("Table5_dw", DigitalProcess.MathTrancate(dw, 2, false));

            // dow
            double dow = hashTable422.GetValue(ptoName, "d0w");
            report.SetValue("Table5_dow", DigitalProcess.MathTrancate(dow, 2, false));

            // условие прочности
            if (dw > dow)
            {
                m_messages.Add("Таблица 5, стр. 13");
            }

            return true;
        }

        private bool FillTable6(ref SReport_Utility.PrintReport report)
        {
            MyDoubleHashTable hashTable423 = m_hashTable423;
            string ptoName = comboBox1.Text;

            // // d0-up d0w-up	Mk-up z-c d0-down d0w-down Мк-down z-down ksi

            // d0 - down
            double d0_down = hashTable423.GetValue(ptoName, "d0down");
            report.SetValue("Table6_d0_down", GetDoubleFormat(d0_down, 0));

            // d0 - up
            double d0_up = hashTable423.GetValue(ptoName, "d0up");
            report.SetValue("Table6_d0_up", GetDoubleFormat(d0_up, 0));

            // dc
            int dc = 0;
            report.SetValue("Table6_dc", dc.ToString());

            // Mk - down
            double Mk_down = hashTable423.GetValue(ptoName, "Мкdown");
            report.SetValue("Table6_Mk_down", GetDoubleFormat(Mk_down, 0));

            // Mk - up
            double Mk_up = hashTable423.GetValue(ptoName, "Mkup");
            report.SetValue("Table6_Mk_up", GetDoubleFormat(Mk_up, 0));

            // ksi
            double ksi = hashTable423.GetValue(ptoName, "ksi");
            report.SetValue("Table6_ksi", DigitalProcess.MathTrancate(ksi, 2, false));

            // Fow - down
            double Fow_down = Mk_down / (ksi * d0_down);
            report.SetValue("Table6_Fow_down", DigitalProcess.MathTrancateInc(Convert.ToInt32(Fow_down), 2).ToString());

            // Fow - up
            double Fow_up = Mk_up / (ksi * d0_up);
            report.SetValue("Table6_Fow_up", DigitalProcess.MathTrancateInc(Convert.ToInt32(Fow_up), 2).ToString());

            // sigma
            SteelProperty steel = m_Krepeg40x;
            double[] X = m_T;
            double[] Y = steel.Rp;

            PhisicsDependency.Lines lines1 = new PhisicsDependency.Lines(X, Y);

            double T2 = 0;
            double.TryParse(textBox11.Text, out T2);

            double sigma = lines1.Value(T2) / 2;
            report.SetValue("Table6_SigmaW", DigitalProcess.MathTrancate(sigma, 1, false));

            // dw - down
            double dw_down = Math.Sqrt(1.27 * Fow_down / sigma + dc * dc);
            report.SetValue("Table6_dw_down", DigitalProcess.MathTrancate(dw_down, 1, false));

            // dw - up
            double dw_up = Math.Sqrt(1.27 * Fow_up / sigma + dc * dc);
            report.SetValue("Table6_dw_up", DigitalProcess.MathTrancate(dw_up, 1, false));

            // d0w - down
            double d0w_down = hashTable423.GetValue(ptoName, "d0wdown");
            report.SetValue("Table6_d0w_down", DigitalProcess.MathTrancate(d0w_down, 1, false));

            // d0w - up
            double d0w_up = hashTable423.GetValue(ptoName, "d0wup");
            report.SetValue("Table6_d0w_up", DigitalProcess.MathTrancate(d0w_up, 1, false));

            // условие прочности
            bool condition = (dw_down < d0w_down) && (dw_up < d0w_up);

            if (!condition)
            {
                m_messages.Add("Таблица 6, стр. 14");
            }

            return true;
        }

        private bool FillTable8(ref SReport_Utility.PrintReport report)
        {
            if (!Is0408())
            {
                MyDoubleHashTable hashTable = m_hashTable424;
                string ptoName = comboBox1.Text;

                // -rib
                // -paronit
                // b0-1	Dm-1	b0-2	Dm-2	b0-3	Dm-3	b1-rib	b2-rib	b3-rib	delta-rib	q0-liq-rib	q0-vap-rib	q0-air-rib	m-liq-rib	m-vap-rib	m-air-rib	nu-rib	b1-paronit	b2-paronit	b3-paronit	delta-paronit	q0-liq-paronit	q0-vap-paronit	q0-air-paronit	m-liq-paronit	m-vap-paronit	m-air-paronit	nu-paronit	hi	ksi	dc	z	d0	d0w

                // ключ для фланцев
                string flan = (comboBox16.SelectedIndex < 3) ? (comboBox16.SelectedIndex + 1).ToString() : "1";

                // ключ для материала прокладок
                string wrapper = (comboBox4.SelectedIndex < 2) ? "paronit" : "rib";

                // b0
                double b0 = hashTable.GetValue(ptoName, "b0" + flan);
                report.SetValue("Table8_b0", DigitalProcess.MathTrancate(b0, 1, false));

                // b1
                double b;

                if (comboBox4.SelectedIndex < 2) // паронит
                {
                    if (b0 > 10)
                    {
                        b = Math.Sqrt(10 * b0);
                    }
                    else
                    {
                        b = b0;
                    }
                }
                else // резина
                {
                    b = b0;
                }

                report.SetValue("Table8_b", DigitalProcess.MathTrancate(b, 1, false));

                // delta
                double delta = hashTable.GetValue(ptoName, "delta" + wrapper);
                report.SetValue("Table8_delta", DigitalProcess.MathTrancate(delta, 1, false));

                // q0
                double q0;
                double q0_Hot, q0_Cold;

                if (comboBox4.SelectedIndex < 2) // паронит
                {
                    // hot
                    if (comboBox13.SelectedIndex == 0) // жидкость
                    {
                        q0_Hot = 80 / Math.Sqrt(10 * delta);
                    }
                    else if (comboBox13.SelectedIndex == 1) // пар
                    {
                        q0_Hot = 100 / Math.Sqrt(10 * delta);
                    }
                    else // газ
                    {
                        q0_Hot = 130 / Math.Sqrt(10 * delta);
                    }

                    // cold
                    if (comboBox14.SelectedIndex == 0) // жидкость
                    {
                        q0_Cold = 80 / Math.Sqrt(10 * delta);
                    }
                    else if (comboBox14.SelectedIndex == 1) // пар
                    {
                        q0_Cold = 100 / Math.Sqrt(10 * delta);
                    }
                    else // газ
                    {
                        q0_Cold = 130 / Math.Sqrt(10 * delta);
                    }
                }
                else // резина
                {
                    // hot
                    if (comboBox13.SelectedIndex == 0) // жидкость
                    {
                        q0_Hot = 5;
                    }
                    else if (comboBox13.SelectedIndex == 1) // пар
                    {
                        q0_Hot = 9;
                    }
                    else // газ
                    {
                        q0_Hot = 13;
                    }

                    // cold
                    if (comboBox14.SelectedIndex == 0) // жидкость
                    {
                        q0_Cold = 5;
                    }
                    else if (comboBox14.SelectedIndex == 1) // пар
                    {
                        q0_Cold = 9;
                    }
                    else // газ
                    {
                        q0_Cold = 13;
                    }
                }

                q0 = Math.Max(q0_Hot, q0_Cold);
                report.SetValue("Table8_q0", DigitalProcess.MathTrancate(q0, 1, false));

                // Dm
                double Dm = hashTable.GetValue(ptoName, "Dm" + flan);
                report.SetValue("Table8_Dm", DigitalProcess.MathTrancate(Dm, 1, false));

                // P
                double P;
                double.TryParse(textBox1.Text, out P);
                report.SetValue("Table8_P", DigitalProcess.MathTrancate(P, 1, true));

                // Ph
                report.SetValue("Table8_PH", DigitalProcess.MathTrancate(m_PressureCalc, 1, true));

                // Fob
                double Fob = Math.PI * Dm * b * q0;
                report.SetValue("Table8_Fob", DigitalProcess.MathTrancate(Convert.ToInt32(Fob), 2).ToString());

                // Fp
                double Fp = Math.PI / 4 * Dm * Dm * P;
                report.SetValue("Table8_Fp", DigitalProcess.MathTrancate(Convert.ToInt32(Fp), 2).ToString());

                // Fd
                double Fd = Fob + 0.1 * Fp;
                report.SetValue("Table8_Fd", DigitalProcess.MathTrancate(Convert.ToInt32(Fd), 2).ToString());

                // Hi
                double hi;
                double.TryParse(report.GetValue("Table5_Hi"), out hi);
                report.SetValue("Table8_Hi", hi.ToString());

                string MediumHot, MediumCold;

                if (comboBox13.SelectedIndex == 0)
                {
                    MediumHot = "liq";
                }
                else if (comboBox13.SelectedIndex == 1)
                {
                    MediumHot = "vap";
                }
                else
                {
                    MediumHot = "air";
                }

                if (comboBox14.SelectedIndex == 0)
                {
                    MediumCold = "liq";
                }
                else if (comboBox14.SelectedIndex == 1)
                {
                    MediumCold = "vap";
                }
                else
                {
                    MediumCold = "air";
                }

                // m
                double m;
                double mHot = hashTable.GetValue(ptoName, "m" + MediumHot + wrapper);
                double mCold = hashTable.GetValue(ptoName, "m" + MediumCold + wrapper);
                m = Math.Max(mHot, mCold);
                report.SetValue("Table8_m", DigitalProcess.MathTrancate(m, 1, false));

                // nu
                double nu = hashTable.GetValue(ptoName, "nu" + wrapper);
                report.SetValue("Table8_eta", DigitalProcess.MathTrancate(nu, 1, false));

                // F2
                double F2 = Math.PI * Dm * b * hi * m * P;
                report.SetValue("Table8_F2", DigitalProcess.MathTrancate(Convert.ToInt32(F2), 2).ToString());

                // F2+
                double F2_plus = F2 + (1 - nu) * Fp;
                report.SetValue("Table8_F2_plus", DigitalProcess.MathTrancate(Convert.ToInt32(F2_plus), 2).ToString());

                // Fph
                double Fph = Math.PI / 4 * Dm * Dm * m_PressureCalc;
                report.SetValue("Table8_Fph", DigitalProcess.MathTrancate(Convert.ToInt32(Fph), 2).ToString());

                // F2h
                double F2h = 0.8 * Math.PI * Dm * b * hi * m * m_PressureCalc;
                report.SetValue("Table8_F2h", DigitalProcess.MathTrancate(Convert.ToInt32(F2h), 2).ToString());

                // F2h_plus
                double F2h_plus = F2h + (1 - nu) * Fph;
                report.SetValue("Table8_F2h_plus", DigitalProcess.MathTrancate(Convert.ToInt32(F2h_plus), 2).ToString());

                // z
                double z = hashTable.GetValue(ptoName, "z");
                report.SetValue("Table8_z", GetDoubleFormat(z, 0));

                // d0
                double d0 = hashTable.GetValue(ptoName, "d0");
                report.SetValue("Table8_d0", GetDoubleFormat(d0, 0));

                // ksi
                double ksi = hashTable.GetValue(ptoName, "ksi");
                report.SetValue("Table8_ksi", DigitalProcess.MathTrancate(ksi, 2, false));

                // Mk
                double max = Fd;

                if (F2_plus > max) max = F2_plus;
                if (F2h_plus > max) max = F2h_plus;

                max *= 1.1;

                double Mk = max * ksi * d0 / z;
                int nMk = DigitalProcess.MathTrancateInc(Convert.ToInt32(Mk), 2);
                report.SetValue("Table8_Mk", nMk.ToString());

                // Fow
                double Fow = nMk * z / (ksi * d0);
                report.SetValue("Table8_Fow", DigitalProcess.MathTrancate(Convert.ToInt32(Fow), 2).ToString());

                // === correct 7 april 2010 ===
                max = Fd;

                if (F2_plus > max) max = F2_plus;
                if (F2h_plus > max) max = F2h_plus;
                if (Fow > max) max = Fow;
                // ============================

                // Fw
                //double Fw = Fow;
                double Fw = max;
                report.SetValue("Table8_Fw", DigitalProcess.MathTrancate(Convert.ToInt32(Fw), 2).ToString());

                // dc
                double dc = hashTable.GetValue(ptoName, "dc");
                report.SetValue("Table8_dc", GetDoubleFormat(dc, 0));

                // dw
                double dw;
                double[] X = m_T;
                SteelProperty steel = m_Krepeg40x;
                double[] Y = steel.Rp;

                PhisicsDependency.Lines lines1 = new PhisicsDependency.Lines(X, Y);

                double T2 = 0;
                double.TryParse(textBox11.Text, out T2);

                double sigma = lines1.Value(T2) / 2;
                dw = Math.Sqrt(1.27 * Fw / (z * sigma) + dc * dc);

                report.SetValue("Table8_dw", DigitalProcess.MathTrancate(dw, 2, false));

                // d0w
                double d0w = hashTable.GetValue(ptoName, "d0w");
                report.SetValue("Table8_d0w", DigitalProcess.MathTrancate(d0w, 2, false));

                // условие прочности
                if (dw > d0w)
                {
                    m_messages.Add("Таблица 8, стр. 17");
                }
            }
            
            return true;
        }

        private bool FillTable9(ref SReport_Utility.PrintReport report)
        {
            MyDoubleHashTable hashTable = m_hashTable43;
            string ptoName = comboBox1.Text;

            // f	n	Am	Bm	A1	B1	b	h	dh	epsi	Epr	q	R	[p]

            // n
            double n = hashTable.GetValue(ptoName, "n");
            report.SetValue("Table9_n", DigitalProcess.MathTrancate(n, 1, false));

            // f
            double f = hashTable.GetValue(ptoName, "f");
            report.SetValue("Table9_f", DigitalProcess.MathTrancate(f, 1, false));

            // h
            double h = hashTable.GetValue(ptoName, "h");
            report.SetValue("Table9_h", DigitalProcess.MathTrancate(h, 1, false));

            // dh
            double dh = h - hashTable.GetValue(ptoName, "dh");
            report.SetValue("Table9_dh", DigitalProcess.MathTrancate(dh, 1, false));
            
            // epsilon
            double epsilon = dh / h;
            report.SetValue("Table9_epsilon", DigitalProcess.MathTrancate(epsilon * 100, 2, false));

            // B
            double B = hashTable.GetValue(ptoName, "Bm");
            report.SetValue("Table9_Bm", DigitalProcess.MathTrancate(B, 1, false));

            // b
            double b = hashTable.GetValue(ptoName, "b");
            report.SetValue("Table9_b", DigitalProcess.MathTrancate(b, 1, false));

            // Epr
            double Epr = 4 * (1 + b / h);
            report.SetValue("Table9_Epr", DigitalProcess.MathTrancate(Epr, 2, false));

            // q
            double q = Epr * epsilon / (1 - epsilon);
            report.SetValue("Table9_q", DigitalProcess.MathTrancate(q, 2, false));

            // A
            double A = hashTable.GetValue(ptoName, "Am");
            report.SetValue("Table9_A", DigitalProcess.MathTrancate(A, 1, false));            

            // A1
            double A1 = hashTable.GetValue(ptoName, "A1");
            report.SetValue("Table9_A1", DigitalProcess.MathTrancate(A1, 1, false));

            // B1
            double B1 = hashTable.GetValue(ptoName, "B1");
            report.SetValue("Table9_B1", DigitalProcess.MathTrancate(B1, 1, false));

            // Bm
            double Bm = hashTable.GetValue(ptoName, "Bm");
            report.SetValue("Table9_Bm", DigitalProcess.MathTrancate(Bm, 1, false));

            // R
            double R = 2 * (A + B) * b * q;
            report.SetValue("Table9_R", DigitalProcess.MathTrancate(Convert.ToInt32(R), 1).ToString());

            // [p]
            double tpt = (f * R) / (n * (A1 + B1) * h * (1 - epsilon));
            report.SetValue("Table9_tpt", DigitalProcess.MathTrancate(tpt, 2, false));

            // Ph
            report.SetValue("Table9_p", DigitalProcess.MathTrancate(m_PressureCalc, 1, false));

            // условие прочности
            if (tpt < m_PressureCalc)
            {
                m_messages.Add("Таблица 9, стр. 18");
            }

            return true;
        }

        private bool FillTable10(ref SReport_Utility.PrintReport report)
        {
            // 41
            //S1, S2

            // 422
            // q0-liq	q0-vap	q0-air	m-liq	m-vap	m-air	eta	hi	z	d0	d0w	ksi

            // 51
            // "k-4", "k-5", "k-6", "k-7", "k-8", "k-9" 

            MyDoubleHashTable hashTable41 = m_hashTablePlit;
            MyDoubleHashTable hashTable422 = m_hashTable422;
            MyDoubleHashTable hashTable43 = m_hashTable43;
            MyDoubleHashTable hashTable51 = m_hashTable51;
            string ptoName = comboBox1.Text;

            double T2;
            double.TryParse(textBox11.Text, out T2);

            // n
            int n;
            int.TryParse(textBox3.Text, out n);
            report.SetValue("Table10_n", n.ToString());

            // s
            double s = m_s[comboBox3.SelectedIndex];
            report.SetValue("Table10_s", s.ToString());

            // l1
            double l1 = n * s;
            report.SetValue("Table10_l1", l1.ToString());

            // dT
            double dT;
            double.TryParse(textBox11.Text, out dT);
            dT -= 20;
            report.SetValue("Table10_dT", dT.ToString());

            // lw
            double s1 = hashTable41.GetValue(ptoName, "S1");
            double s2 = hashTable41.GetValue(ptoName, "S2");
            double k = hashTable51.GetValue(ptoName, "k" + (comboBox3.SelectedIndex + 4).ToString());

            if (k == 0)
            {
                m_lastError = $"Расчет с толщиной пластины {m_s[comboBox3.SelectedIndex]} мм не предусмотрен для {ptoName}.";
                return false;
            }

            double lw = n * k + s1 + s2;
            report.SetValue("Table10_lw", DigitalProcess.MathTrancate(lw, 1, false));

            // d
            double d = hashTable422.GetValue(ptoName, "d0");
            report.SetValue("Table10_d", d.ToString());

            // z
            double z = hashTable422.GetValue(ptoName, "z");
            report.SetValue("Table10_z", z.ToString());            

            double[] X = m_T;
            double[] Y1 = { 215000, 212000, 210000, 207000, 205000, 202000 };

            PhisicsDependency.Lines lines = new PhisicsDependency.Lines(X, Y1);

            // Ew
            double Ew = lines.Value(T2);
            report.SetValue("Table10_Ew", GetDoubleFormat(Ew, 0));

            // dw
            double dw = hashTable422.GetValue(ptoName, "d0w");
            report.SetValue("Table10_dw", dw.ToString());

            // Aw
            double Aw = 0.785 * dw * dw;
            report.SetValue("Table10_Aw", GetDoubleFormat(Aw, 0));

            // LambdaW
            double LambdaW = (lw + 0.6 * d) / (z * Ew * Aw);
            report.SetValue("Table10_LambdaW", DigitalProcess.MathTrancateToDigital(LambdaW, 9));

            // alpha1
            double[] Y2 = { 0, 11.5, 11.9, 12.2, 12.5, 12.8 };
            lines = new PhisicsDependency.Lines(X, Y2);

            double Alpha1 = lines.Value(T2);
            Alpha1 *= 0.000001;
            report.SetValue("Table10_Alpha1", DigitalProcess.MathTrancateToDigital(Alpha1, 9));

            // alpha2
            double[] Y3 = (comboBox2.SelectedIndex == 0) ? new double[] { 0, 16.4, 16.6, 16.8, 17, 17.2 } : new double[] { 0, 7.8, 7.8, 8, 8.3, 8.5 };

            lines = new PhisicsDependency.Lines(X, Y3);

            double Alpha2 = lines.Value(T2) * 0.000001;
            report.SetValue("Table10_Alpha2", DigitalProcess.MathTrancateToDigital(Alpha2, 9));

            // eta
            double eta = hashTable422.GetValue(ptoName, "eta");
            report.SetValue("Table10_eta", DigitalProcess.MathTrancate(eta, 1, false));

            // ================ Mk =====================================================

            // A
            double A = hashTable43.GetValue(ptoName, "Am");

            // B
            double B = hashTable43.GetValue(ptoName, "Bm");

            // b
            double b = hashTable43.GetValue(ptoName, "b");

            // q0
            int MediumHotIndex = comboBox13.SelectedIndex;
            int MediumColdIndex = comboBox14.SelectedIndex;
            double q0_Hot, q0_Cold;
            string sQuery;

            if (MediumHotIndex == 0) // liq
            {
                sQuery = "q0liq";
            }
            else if (MediumHotIndex == 1) // vap
            {
                sQuery = "q0vap";
            }
            else // air
            {
                sQuery = "q0air";
            }

            q0_Hot = hashTable422.GetValue(ptoName, sQuery);

            if (MediumColdIndex == 0) // liq
            {
                sQuery = "q0liq";
            }
            else if (MediumColdIndex == 1) // vap
            {
                sQuery = "q0vap";
            }
            else // air
            {
                sQuery = "q0air";
            }

            q0_Cold = hashTable422.GetValue(ptoName, sQuery);

            double q0 = Math.Max(q0_Hot, q0_Cold);

            // Fob
            double Fob = 2 * (A + B) * b * q0;

            // P
            double P;
            double.TryParse(textBox1.Text, out P);

            // Fp
            double Fp = A * B * P;

            // Fd
            double Fd = Fob + 0.1 * Fp;

            // Hi
            double[] X1 = { 20, 200 };
            double[] Y4 = { 1.0, 1.5 };

            lines = new PhisicsDependency.Lines(X1, Y4);
            double Hi = lines.Value(T2);

            // m
            double m_Hot, m_Cold;

            if (MediumHotIndex == 0) // liq
            {
                sQuery = "mliq";
            }
            else if (MediumHotIndex == 1) // vap
            {
                sQuery = "mvap";
            }
            else // air
            {
                sQuery = "mair";
            }

            m_Hot = hashTable422.GetValue(ptoName, sQuery);

            if (MediumColdIndex == 0) // liq
            {
                sQuery = "mliq";
            }
            else if (MediumColdIndex == 1) // vap
            {
                sQuery = "mvap";
            }
            else // air
            {
                sQuery = "mair";
            }

            m_Cold = hashTable422.GetValue(ptoName, sQuery);

            double m = Math.Max(m_Hot, m_Cold);

            // F2
            double F2 = 2 * (A + B) * b * Hi * m * P;

            // F2 plus
            double F2_plus = F2 + (1 - eta) * Fp;

            // Fph
            double Fph = A * B * m_PressureCalc;

            // F2h
            double F2h = 0.8 * 2 * (A + B) * b * Hi * m * m_PressureCalc;

            // F2h_plus
            double F2h_plus = F2h + (1 - eta) * Fph;

            // d0
            double d0 = hashTable422.GetValue(ptoName, "d0");

            // ksi
            double ksi = hashTable422.GetValue(ptoName, "ksi");

            // Mk
            double max = Fd;

            if (F2_plus > max) max = F2_plus;
            if (F2h_plus > max) max = F2h_plus;

            max *= 1.1;

            double Mk = max * ksi * d0 / z;
            int nMk = DigitalProcess.MathTrancateInc(Convert.ToInt32(Mk), 2);
            report.SetValue("Table10_Mk", nMk.ToString());

            // =========================================================================

            // Fw
            max = Fd;

            if (F2_plus > max) max = F2_plus;
            if (F2h_plus > max) max = F2h_plus;

            double Fow = Mk * z / (ksi * d);

            if (Fow > max) max = Fow;

            double Fw = max;

            report.SetValue("Table10_Fw", DigitalProcess.MathTrancate(Convert.ToInt32(Fw), 2).ToString());

            // Fp-1
            report.SetValue("Table10_Fp_1", "-");

            // Fp-2
            report.SetValue("Table10_Fp_2", DigitalProcess.MathTrancate(Convert.ToInt32(Fp), 2).ToString());

            // Fp-3
            report.SetValue("Table10_Fp_3", DigitalProcess.MathTrancate(Convert.ToInt32(Fp), 2).ToString());

            // Fp-4
            report.SetValue("Table10_Fp_4", "-");

            // Ft-1
            report.SetValue("Table10_Ft_1", "-");

            // Ft-2
            report.SetValue("Table10_Ft_2", "-");

            // Ft-3
            double Ft = Math.Abs(Alpha2 - Alpha1) * l1 * dT / LambdaW;
            report.SetValue("Table10_Ft_3", DigitalProcess.MathTrancate(Convert.ToInt32(Ft), 2).ToString());

            // Ft-4
            report.SetValue("Table10_Ft_4", "-");

            // Fph-1
            report.SetValue("Table10_Fph_1", "-");

            // Fph-2
            report.SetValue("Table10_Fph_2", "-");

            // Fph-3
            report.SetValue("Table10_Fph_3", "-");

            // Fph-4
            report.SetValue("Table10_Fph_4", DigitalProcess.MathTrancate(Convert.ToInt32(Fph), 2).ToString());

            // Fw-1
            double Fwp1 = Fw;
            report.SetValue("Table10_Fwp_1", DigitalProcess.MathTrancate(Convert.ToInt32(Fwp1), 2).ToString());

            // Fw-2
            double Fwp2 = Fw + eta * Fp;
            report.SetValue("Table10_Fwp_2", DigitalProcess.MathTrancate(Convert.ToInt32(Fwp2), 2).ToString());

            // Fw-3
            double Fwp3 = Fwp2 + Ft;
            report.SetValue("Table10_Fwp_3", DigitalProcess.MathTrancate(Convert.ToInt32(Fwp3), 2).ToString());

            // Fw-4
            double Fwp4 = Fw + eta * A * B * m_PressureCalc;
            report.SetValue("Table10_Fwp_4", DigitalProcess.MathTrancate(Convert.ToInt32(Fwp4), 2).ToString());

            // Sigma round
            double SigmaRound1 = Fwp1 / (Aw * z);
            double SigmaRound2 = Fwp2 / (Aw * z);
            double SigmaRound3 = Fwp3 / (Aw * z);
            double SigmaRound4 = Fwp4 / (Aw * z);

            report.SetValue("Table10_SigmaW_Round_1", DigitalProcess.MathTrancate(SigmaRound1, 1, false));
            report.SetValue("Table10_SigmaW_Round_2", DigitalProcess.MathTrancate(SigmaRound2, 1, false));
            report.SetValue("Table10_SigmaW_Round_3", DigitalProcess.MathTrancate(SigmaRound3, 1, false));
            report.SetValue("Table10_SigmaW_Round_4", DigitalProcess.MathTrancate(SigmaRound4, 1, false));

            // sigma
            SteelProperty steel = m_Krepeg40x;
            double[] Y5 = steel.Rp;
            lines = new PhisicsDependency.Lines(X, Y5);
            double Sigma = lines.Value(T2) / 2;

            double SigmaSquareW1 = Sigma;
            double SigmaSquareW2 = Sigma;
            double SigmaSquareW3 = 1.3 * Sigma;
            double SigmaSquareW4 = 1.6 * Sigma;

            report.SetValue("Table10_SigmaW_Square_1", DigitalProcess.MathTrancate(SigmaSquareW1, 1, false));
            report.SetValue("Table10_SigmaW_Square_2", DigitalProcess.MathTrancate(SigmaSquareW2, 1, false));
            report.SetValue("Table10_SigmaW_Square_3", DigitalProcess.MathTrancate(SigmaSquareW3, 1, false));
            report.SetValue("Table10_SigmaW_Square_4", DigitalProcess.MathTrancate(SigmaSquareW4, 1, false));

            // tau
            double tau = Mk / (2 * 0.196 * dw * dw * dw);
            report.SetValue("Table10_Tau", DigitalProcess.MathTrancate(tau, 1, false));

            // Sigma round 4
            double SigmaRound4W1 = Math.Sqrt(SigmaRound1 * SigmaRound1 + 4 * tau * tau);
            double SigmaRound4W2 = Math.Sqrt(SigmaRound2 * SigmaRound2 + 4 * tau * tau);
            double SigmaRound4W3 = Math.Sqrt(SigmaRound3 * SigmaRound3 + 4 * tau * tau);
            double SigmaRound4W4 = Math.Sqrt(SigmaRound4 * SigmaRound4 + 4 * tau * tau);

            report.SetValue("Table10_Sigma4W_Round_1", DigitalProcess.MathTrancate(SigmaRound4W1, 1, false));
            report.SetValue("Table10_Sigma4W_Round_2", DigitalProcess.MathTrancate(SigmaRound4W2, 1, false));
            report.SetValue("Table10_Sigma4W_Round_3", DigitalProcess.MathTrancate(SigmaRound4W3, 1, false));
            report.SetValue("Table10_Sigma4W_Round_4", DigitalProcess.MathTrancate(SigmaRound4W4, 1, false));

            // sigma square 4
            double SigmaSquare4W1 = Sigma;
            double SigmaSquare4W2 = 1.7 * Sigma;
            double SigmaSquare4W3 = 1.7 * Sigma;
            double SigmaSquare4W4 = 1.8 * Sigma;

            report.SetValue("Table10_Sigma4W_Square_1", DigitalProcess.MathTrancate(SigmaSquare4W1, 1, false));
            report.SetValue("Table10_Sigma4W_Square_2", DigitalProcess.MathTrancate(SigmaSquare4W2, 1, false));
            report.SetValue("Table10_Sigma4W_Square_3", DigitalProcess.MathTrancate(SigmaSquare4W3, 1, false));
            report.SetValue("Table10_Sigma4W_Square_4", DigitalProcess.MathTrancate(SigmaSquare4W4, 1, false));

            // условие прочности
            bool Condition = SigmaRound1 < SigmaSquareW1;
            Condition = Condition && (SigmaRound2 < SigmaSquareW2);
            Condition = Condition && (SigmaRound3 < SigmaSquareW3);
            Condition = Condition && (SigmaRound4 < SigmaSquareW4);

            Condition = Condition && (SigmaRound4W1 < SigmaSquare4W1);
            Condition = Condition && (SigmaRound4W2 < SigmaSquare4W2);
            Condition = Condition && (SigmaRound4W3 < SigmaSquare4W3);
            Condition = Condition && (SigmaRound4W4 < SigmaSquare4W4);

            if (!Condition)
            {
                m_messages.Add("Таблица 10, стр. 21");
            }

            Bitmap scBitmap = null;

            // Рисунок 3 (сечение)

            // 04-08
            Set set0408 = m_set_0408;

            Set setSmall1 = new Set(12, 23);
            Set setSmall2 = new Set(24, 47);            
            Set setSmall3 = new Set(154, 157);

            // 41
            Set set41 = new Set(48, 53);            

            // 42
            Set set42 = new Set(54, 59);

            // 62
            Set set62 = new Set(60, 65);

            // 86
            Set set86 = new Set(66, 71);

            // 110
            Set set110 = new Set(72, 77);

            // 43
            Set set43 = new Set(78, 83);

            // 65
            Set set65 = new Set(84, 89);

            // 100
            Set set100 = new Set(90, 95);

            // 130
            Set set130 = new Set(96, 101);

            // 152
            Set set152 = new Set(102, 107);

            // 220
            Set set220 = new Set(90, 95);

            // 113
            Set set113 = new Set(114, 119);

            // 81
            Set set81 = new Set(120, 125);

            // 121
            Set set121 = new Set(126, 131);

            // 188
            Set set188 = new Set(132, 137);

            // 251
            Set set251 = new Set(138, 143);

            // 145
            Set set145 = new Set(144, 147);

            // 210
            Set set210 = new Set(148, 151);

            // 53
            Set set53 = new Set(158, 159);

            // 160
            Set set160 = new Set(160, 161);

            // 229
            Set set229 = new Set(179, 180);

            // 26
            Set set26 = new Set(170, 175);

            // 31
            Set set31 = new Set(176, 178);

            // 19w
            Set set19w = new Set(162, 169);

            int ptoIndex = comboBox1.SelectedIndex;
            
            if (set0408.In(ptoIndex))
            {
                scBitmap = SimpleImageProccessing.Zoom(Properties.Resources.napr_0408, 620, 620);
            }
            else if (setSmall1.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Sech_07_20, 650, 650);
            }
            else if (setSmall2.In(ptoIndex) || set26.In(ptoIndex))
            {
                //scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Sech_21_47, 1021, 1021);
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.sech_21_47_new, 294, 294);
            }
            else if (setSmall3.In(ptoIndex) || set19w.In(ptoIndex) || set31.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.sech_19, 650, 650);
            }
            else if (set41.In(ptoIndex) || set53.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._41_42_62_86__110_, 639, 639);
            }
            else if (set42.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._41_42_62_86__110_, 639, 639);
            }
            else if (set62.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._41_42_62_86__110_, 639, 639);
            }
            else if (set86.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._41_42_62_86__110_, 639, 639);
            }
            else if (set110.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Kozenkova_luba, 648, 648);
            }
            else if (set43.In(ptoIndex) || set160.In(ptoIndex) || set229.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Kozenkova_luba, 648, 648);
            }
            else if (set65.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Kozenkova_luba, 648, 648);
            }
            else if (set100.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Kozenkova_luba, 648, 648);
            }
            else if (set130.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 2000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._1_big, 741, 741);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._2_big, 808, 808);
                }
            }
            else if (set152.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._3_big, 776, 776);
            }
            else if (set220.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L < 4000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._2_big, 808, 808);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._6_big, 750, 750);
                }
            }
            else if (set113.In(ptoIndex))
            {
                scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Kozenkova_luba, 648, 648);
            }
            else if (set81.In(ptoIndex))
            {
                int frameIndex = comboBox17.SelectedIndex;

                if (frameIndex > 6)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._81_8_9_10, 626, 626);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Kozenkova_luba, 648, 648);
                }
            }
            else if (set121.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 2000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._1_big, 741, 741);
                }
                else if (L <= 4000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._2_big, 808, 808);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._6_big, 750, 750);
                }
            }
            else if (set188.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 2000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._1_big, 741, 741);
                }
                else if (L <= 4000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._2_big, 808, 808);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._4_big, 742, 742);
                }
            }
            else if (set251.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 2000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._1_big, 741, 741);
                }
                else if (L <= 4000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._6_big, 750, 750);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._4_big, 742, 742);
                }
            }
            else if (set145.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 2000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._1_big, 741, 741);
                }
                else if (L <= 4000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._2_big, 808, 808);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._4_big, 742, 742);
                }
            }
            else if (set210.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 4000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._2_big, 808, 808);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._4_big, 742, 742);
                }
            }
            else // if (set201.In(ptoIndex))
            {
                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 4000)
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._6_big, 750, 750);
                }
                else
                {
                    scBitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources._4_big, 742, 742);
                }
            }

            report.SetPicture("Picture3", scBitmap);

            return true;
        }

        private bool FillTable11(ref SReport_Utility.PrintReport report)
        {
            // picture 4
            Bitmap picture4;

            if (!IsBig())
            {
                if (Is212247())
                {
                    picture4 = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Schema1, 2956, 947);
                }
                else
                {
                    picture4 = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Picture4, 2956, 947);
                }
            }
            else
            {
                picture4 = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.pic4_big, 765, 331);
            }
            
            report.SetPicture("Picture4", picture4);

            // picture 5
            Bitmap picture5 = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.Picture5, 2956, 947);
            report.SetPicture("Picture5", picture5);

            int N;
            int.TryParse(textBox3.Text, out N);

            MyDoubleHashTable hashTable51 = m_hashTable51;
            string ptoName = comboBox1.Text;

            // L1
            double L1;
            double k = hashTable51.GetValue(ptoName, "k" + (comboBox3.SelectedIndex + 4).ToString());

            if (k == 0)
            {
                m_lastError = $"Расчет с толщиной пластины {m_s[comboBox3.SelectedIndex]} мм не предусмотрен для {ptoName}.";
                return false;
            }

            L1 = N * k;
            report.SetValue("Table11_L1_1", L1.ToString());

            // L2
            double L, L2;
            double.TryParse(comboBox17.Text, out L);
            L2 = L - L1;
            report.SetValue("Table11_L2_1", L2.ToString());

            // L3
            report.SetValue("Table11_L3_1", L.ToString());

            // Q - 1
            report.SetValue("Table11_Q_1", "-");

            // Q - 2
            double Q;

            // масса пустой
            double MNetto;
            double.TryParse(textBox10.Text, out MNetto);

            // масса заполненный
            double MBrutto;
            double.TryParse(textBox9.Text, out MBrutto);

            // масса среды
            double Msh = MBrutto - MNetto;

            string sQuery;

            // Масса прокладки
            string wrapper = comboBoxWrapper.Text;
            sQuery = wrapper.ToLower();
            double Mwr = m_list52.GetValue(ptoName, sQuery);

            // Масса пластины
            if (comboBox2.SelectedIndex == 0)
            {
                sQuery = "aisi";
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                sQuery = "aisi";
            }
            else
            {
                sQuery = "titan5";
            }

            sQuery += "m" + (comboBox3.SelectedIndex + 4).ToString();
            double Mpl = m_list52.GetValue(ptoName, sQuery);

            Q = (Mwr + Mpl) * N + Msh;
            Q *= 10;
            Q /= L1;

            report.SetValue("Table11_Q_2", GetDoubleFormat(Q, 1));

            // F - 1
            double F = m_list52.GetValue(ptoName, "F");
            F *= 10;
            report.SetValue("Table11_F_1", DigitalProcess.MathTrancate(Convert.ToInt32(F), 1).ToString());

            // F - 2
            report.SetValue("Table11_F_2", "-");

            // M - 1
            double M1, M2;

            if (IsBig())
            {
                // реакция опоры А
                double Ra = F * (1 - L1 / L) + Q * L1 * (1 - L1 / (2 * L));

                // picture = formula page 23
                Bitmap bitmap1 = null;

                if (Ra / Q < L1)
                {
                    M1 = (Ra * Ra) / (2 * Q);
                    bitmap1 = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.M2, 230, 109);
                }
                else
                {
                    M1 = L1 * (Ra - Q * L1 / 2);
                    bitmap1 = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.M1, 230, 109);
                }

                M2 = 0;

                report.SetPicture("Picture9", bitmap1);
                report.SetValue("Table11_Ra", GetDoubleFormat(Ra, 0));
            }
            else
            {
                // M1
                M1 = Q / 2;
                M1 *= L1 * L1 * (2 * L2 + L1) * (2 * L2 + L1) / (4 * L * L);

                // M2
                M2 = F * L1 * L2 / L;

                //if (Is212247())
                //{
                //    M1 += M2;
                //}
            }

            report.SetValue("Table11_M_1", DigitalProcess.MathTrancate(Convert.ToInt32(M1 + M2), 1).ToString());
            report.SetValue("Table11_M_2", DigitalProcess.MathTrancate(Convert.ToInt32(M2), 1).ToString());

            // W - 1
            sQuery = "W1" + comboBox17.Text;
            double W1 = m_len52.GetValue(ptoName, sQuery);
            report.SetValue("Table11_W_1", W1.ToString());

            // W - 2
            double W2 = m_len52.GetValue(ptoName, "W2");
            report.SetValue("Table11_W_2", W2.ToString());

            // Sigma - 1
            double Sigma1 = M1 / W1;
            double Sigma2 = M2 / W1;

            if (Is212247())
            {
                Sigma1 += Sigma2;
            }

            report.SetValue("Table11_Sigma_1", GetDoubleFormat(Sigma1, 2));
            report.SetValue("Table11_Sigma_2", GetDoubleFormat(Sigma2, 2));

            // Направляющая верхняя
            double SigmaSquare1, SigmaSquare2;

            string sVal1 = report.GetValue("Table2_Sigma_22");
            double.TryParse(sVal1, out SigmaSquare1);

            // Направляющая нижняя
            string sVal2 = report.GetValue("Table2_Sigma_32");
            double.TryParse(sVal2, out SigmaSquare2);

            SigmaSquare1 *= 1.3;
            SigmaSquare2 *= 1.3;

            // 13Sigma - 1
            report.SetValue("Table11_13Sigma_1", SigmaSquare1.ToString());

            // 13Sigma - 2
            report.SetValue("Table11_13Sigma_2", SigmaSquare2.ToString());

            // условие прочности
            bool Condition = Sigma1 < SigmaSquare1;
            Condition = Condition && (Sigma2 < SigmaSquare2);

            if (!Condition)
            {
                m_messages.Add("Таблица 11, стр. 24");
            }

            // фраза 5.3.1
            string text1 = "5.3.1 Болты 15, 16 (болты 13, 14 не рассматриваются, т.к. менее нагружены) ";
            text1 += "крепления направляющей верхней к стойке и плите неподвижной ";
            text1 += "испытывают напряжение растяжения от усилий затяга, напряжение кручения от ";
            text1 += "момента затяга";

            string text2;

            if (m_bNaprIn)
            {
                text2 = " и напряжение среза от весовой нагрузки плиты прижимной.";
            }
            else
            {
                text2 = ". Болты на срез не работают, т.к. направляющая устанавливается ";
                text2 += "непосредственно на плиту прижимную и стойку заднюю.";
            }

            report.SetValue("Table11_Phrase1", text1 + text2);

            if (IsBig())
            {
                int Ls = 33;
                double t = 5.2;

                Q = (Mwr + Mpl) * N + Msh;
                Q *= 10;

                double sigma = Q / (L1 * t);
                double h = 8.2; // for all pto from HH41 etc.

                report.SetValue("Table11_Q", GetDoubleFormat(Q, 1));
                report.SetValue("Table11_L", Ls.ToString());
                report.SetValue("Table11_t", t.ToString());
                report.SetValue("Table11_Sigma", DigitalProcess.MathTrancate(sigma, 1, false));
                report.SetValue("Table11_h", h.ToString());

                //  сигма изгиб
                sigma = (6 * Q * Ls) / (2 * L1 * h * h);
                report.SetValue("Table11_Sigma_isgib", DigitalProcess.MathTrancate(sigma, 1, false));
            }

            return true;
        }

        private bool FillTable12(ref SReport_Utility.PrintReport report)
        {
            string sPrefix = "Table12_";
            string ptoName = comboBox1.Text;

            int N;
            int.TryParse(textBox3.Text, out N);

            // F

            // Fow - down
            double Fdown;
            double.TryParse(report.GetValue("Table6_Fow_down"), out Fdown);
            report.SetValue(sPrefix + "F_1314", report.GetValue("Table6_Fow_down"));

            // Fow - up
            double Fup;
            double.TryParse(report.GetValue("Table6_Fow_up"), out Fup);
            report.SetValue(sPrefix + "F_1516", report.GetValue("Table6_Fow_up"));

            // d0w - down
            double d0w_down = m_hashTable423.GetValue(ptoName, "d0wdown");
            double Adown = Math.PI / 4 * d0w_down * d0w_down;
            report.SetValue(sPrefix + "A_1314", DigitalProcess.MathTrancate(Adown, 1, false));

            // d0w - up
            double d0w_up = m_hashTable423.GetValue(ptoName, "d0wup");
            double Aup = Math.PI / 4 * d0w_up * d0w_up;
            report.SetValue(sPrefix + "A_1516", DigitalProcess.MathTrancate(Aup, 1, false));

            // Sigma
            double SigmaDown = Fdown / Adown;
            report.SetValue(sPrefix + "Sigma_1314", DigitalProcess.MathTrancate(SigmaDown, 1, false));

            double SigmaUp = Fup / Aup;
            report.SetValue(sPrefix + "Sigma_1516", DigitalProcess.MathTrancate(SigmaUp, 1, false));

            // Sigma square
            double SigmaSquare;
            double.TryParse(report.GetValue("Table2_Sigma_62"), out SigmaSquare);

            report.SetValue(sPrefix + "SigmaSquare_1314", SigmaSquare.ToString());
            report.SetValue(sPrefix + "SigmaSquare_1516", SigmaSquare.ToString());

            bool Condition = (SigmaDown < SigmaSquare);
            Condition = Condition && (SigmaUp < SigmaSquare);

            // Q - 2
            double Q1;

            // масса пустой
            double MNetto;
            double.TryParse(textBox10.Text, out MNetto);

            // масса заполненный
            double MBrutto;
            double.TryParse(textBox9.Text, out MBrutto);

            // масса среды
            double Msh = MBrutto - MNetto;

            // Масса прокладки
            string wrapper = comboBoxWrapper.Text;
            string sQuery = wrapper.ToLower();
            double Mwr = m_list52.GetValue(ptoName, sQuery);

            // Масса пластины
            if (comboBox2.SelectedIndex == 0)
            {
                sQuery = "aisi";
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                sQuery = "aisi";
            }
            else
            {
                sQuery = "titan5";
            }

            sQuery += "m" + (comboBox3.SelectedIndex + 4).ToString();
            double Mpl = m_list52.GetValue(ptoName, sQuery);

            Q1 = (Mwr + Mpl) * N + Msh;
            Q1 *= 10;

            report.SetValue(sPrefix + "Q1_1314", DigitalProcess.MathTrancate(Q1, 1, false));
            report.SetValue(sPrefix + "Q1_1516", report.GetValue("Table11_F_1"));

            // mu
            double mu = m_list53.GetValue(ptoName, "mu");
            report.SetValue(sPrefix + "mu_1314", mu.ToString());
            report.SetValue(sPrefix + "mu_1516", mu.ToString());

            // z
            double z = m_hashTable423.GetValue(ptoName, "zc");
            z *= 2;
            report.SetValue(sPrefix + "z_1314", z.ToString());
            report.SetValue(sPrefix + "z_1516", z.ToString());

            // Ft
            double Ftdown = mu * Fdown * z;
            double Ftup = mu * Fup * z;
            report.SetValue(sPrefix + "Ft_1314", DigitalProcess.MathTrancate(Ftdown, 1, false));
            report.SetValue(sPrefix + "Ft_1516", DigitalProcess.MathTrancate(Ftup, 1, false));

            // Q
            double F;
            double.TryParse(report.GetValue("Table11_F_1"), out F);
            double Qdown = Q1 - Ftdown;
            double Qup = F - Ftup;
            report.SetValue(sPrefix + "Q_1314", DigitalProcess.MathTrancate(Qdown, 0, false));
            report.SetValue(sPrefix + "Q_1516", DigitalProcess.MathTrancate(Qup, 0, false));

            // d0
            report.SetValue(sPrefix + "d0_1314", DigitalProcess.MathTrancate(d0w_down, 2, false));
            report.SetValue(sPrefix + "d0_1516", DigitalProcess.MathTrancate(d0w_up, 2, false));            

            // tau
            double TauDown = (m_bNaprIn) ? (6 * Qdown / (Math.PI * z * d0w_down * d0w_down)) : 0;
            double TauUp = (m_bNaprIn) ? (6 * Qup / (Math.PI * z * d0w_up * d0w_up)) : 0;

            report.SetValue(sPrefix + "tau_1314", DigitalProcess.MathTrancate(TauDown, 1, false));
            report.SetValue(sPrefix + "tau_1516", DigitalProcess.MathTrancate(TauUp, 1, false));

            if (TauDown < 0) TauDown = 0;
            if (TauUp < 0) TauUp = 0;

            // tau square
            double TauSquare = SigmaSquare / 2;

            report.SetValue(sPrefix + "tauSquare_1314", DigitalProcess.MathTrancate(TauSquare, 1, false));
            report.SetValue(sPrefix + "tauSquare_1516", DigitalProcess.MathTrancate(TauSquare, 1, false));

            //Condition = Condition && (TauDown < TauSquare);
            //Condition = Condition && (TauUp < TauSquare);

            // Mk
            double Mkdown = m_hashTable423.GetValue(ptoName, "Мкdown");
            double Mkup = m_hashTable423.GetValue(ptoName, "Mkup");

            report.SetValue(sPrefix + "Mk_1314", DigitalProcess.MathTrancate(Mkdown, 0, false));
            report.SetValue(sPrefix + "Mk_1516", DigitalProcess.MathTrancate(Mkup, 0, false));

            // Tau R
            double TauKdown = Mkdown / (0.392 * d0w_down * d0w_down * d0w_down);
            double TauKup = Mkup / (0.392 * d0w_up * d0w_up * d0w_up);

            report.SetValue(sPrefix + "tauK_1314", DigitalProcess.MathTrancate(TauKdown, 1, false));
            report.SetValue(sPrefix + "tauK_1516", DigitalProcess.MathTrancate(TauKup, 1, false));

            // sigma 4w
            double Sigma4Wdown = Math.Sqrt(SigmaDown * SigmaDown + 4 * (TauDown + TauKdown) * (TauDown + TauKdown));
            double Sigma4Wup = Math.Sqrt(SigmaUp * SigmaUp + 4 * (TauUp + TauKup) * (TauUp + TauKup));

            report.SetValue(sPrefix + "Sigma4w_1314", DigitalProcess.MathTrancate(Sigma4Wdown, 1, false));
            report.SetValue(sPrefix + "Sigma4w_1516", DigitalProcess.MathTrancate(Sigma4Wup, 1, false));

            // 1.7 * Sigma
            report.SetValue(sPrefix + "17Sigma_1314", DigitalProcess.MathTrancate(1.7 * SigmaSquare, 1, false));
            report.SetValue(sPrefix + "17Sigma_1516", DigitalProcess.MathTrancate(1.7 * SigmaSquare, 1, false));

            Condition = Condition && (Sigma4Wdown < 1.7 * SigmaSquare);
            Condition = Condition && (Sigma4Wup < 1.7 * SigmaSquare);

            // условие прочности
            if (!Condition)
            {
                m_messages.Add("Таблица 12, стр 25, 26");
            }

            // picture = tauK, page 25
            Bitmap bitmap1 = SimpleImageProccessing.Zoom(Properties.Resources.TauK_25, 508, 208);

            report.SetPicture("Picture8", bitmap1);
            
            return true;
        }

        private bool FillTable13(ref SReport_Utility.PrintReport report)
        {
            string ptoName = comboBox1.Text;
            int i;

            // F
            report.SetValue("Table13_F_1", report.GetValue("Table8_Fw"));
            report.SetValue("Table13_F_2", report.GetValue("Table8_Fw"));

            report.SetValue("Table13_F_3", report.GetValue("Table6_Fow_down"));
            report.SetValue("Table13_F_4", report.GetValue("Table6_Fow_down"));
            report.SetValue("Table13_F_5", report.GetValue("Table6_Fow_up"));
            report.SetValue("Table13_F_6", report.GetValue("Table6_Fow_up"));

            report.SetValue("Table13_F_7", report.GetValue("Table10_Fwp_3"));
            report.SetValue("Table13_F_8", report.GetValue("Table10_Fwp_3"));

            // d
            report.SetValue("Table13_d_1", report.GetValue("Table8_d0"));
            report.SetValue("Table13_d_2", report.GetValue("Table8_d0"));

            report.SetValue("Table13_d_3", m_hashTable423.GetValue(ptoName, "d0down").ToString());
            report.SetValue("Table13_d_4", m_hashTable423.GetValue(ptoName, "d0down").ToString());
            report.SetValue("Table13_d_5", m_hashTable423.GetValue(ptoName, "d0up").ToString());
            report.SetValue("Table13_d_6", m_hashTable423.GetValue(ptoName, "d0up").ToString());

            report.SetValue("Table13_d_7", m_hashTable422.GetValue(ptoName, "d0").ToString());
            report.SetValue("Table13_d_8", m_hashTable422.GetValue(ptoName, "d0").ToString());
            
            // h
            for (i = 1; i <= 8; i++)
            {
                report.SetValue("Table13_h_" + i.ToString(), m_list54.GetValue(ptoName, "h" + i.ToString()).ToString());
            }

            // K1
            for (i = 0; i < 8; i++)
            {
                report.SetValue("Table13_K1_" + (i + 1).ToString(), m_list54.GetValue(ptoName, "K1" + ((i % 2 == 0) ? "sh" : "g")).ToString());
            }

            // Km
            double Km = m_list54.GetValue(ptoName, "Km");
            report.SetValue("Table13_Km", Km.ToString());

            // z
            report.SetValue("Table13_z_1", m_hashTable424.GetValue(ptoName, "z").ToString());
            report.SetValue("Table13_z_2", m_hashTable424.GetValue(ptoName, "z").ToString());

            int zc = 1;

            report.SetValue("Table13_z_3", zc.ToString());
            report.SetValue("Table13_z_4", zc.ToString());
            report.SetValue("Table13_z_5", zc.ToString());
            report.SetValue("Table13_z_6", zc.ToString());

            report.SetValue("Table13_z_7", m_hashTable422.GetValue(ptoName, "z").ToString());
            report.SetValue("Table13_z_8", m_hashTable422.GetValue(ptoName, "z").ToString());

            // tau
            double tau;
            double Fw, z, d, h, K1;
            double[] TauP = new double[8];

            for (i = 1; i <= 8; i++)
            {
                // Fw
                double.TryParse(report.GetValue("Table13_F_" + i.ToString()), out Fw);

                // z
                double.TryParse(report.GetValue("Table13_z_" + i.ToString()), out z);

                // d
                double.TryParse(report.GetValue("Table13_d_" + i.ToString()), out d);

                // h
                double.TryParse(report.GetValue("Table13_h_" + i.ToString()), out h);

                // K1
                double.TryParse(report.GetValue("Table13_K1_" + i.ToString()), out K1);

                tau = Fw / (Math.PI * z * d * h * K1 * Km);
                TauP[i - 1] = tau;

                report.SetValue("Table13_TauP_" + i.ToString(), DigitalProcess.MathTrancate(tau, 1, false));
            }

            // tau square
            double[] TauSquareP = new double[8];

            // Rp0,2 болт
            double Rp02bolt;
            double.TryParse(report.GetValue("Table2_PredelTech_62"), out Rp02bolt);
            Rp02bolt /= 4;

            // гайка
            double Rp02gaika;
            double.TryParse(report.GetValue("Table2_PredelTech_72"), out Rp02gaika);
            Rp02gaika /= 4;

            // плита неподвижная
            double Rp02plita;
            double.TryParse(report.GetValue("Table2_PredelTech_12"), out Rp02plita);
            Rp02plita /= 4;

            // резьбовая часть в направляющих
            double Rp02napr;
            double.TryParse(report.GetValue("Table2_PredelTech_42"), out Rp02napr);
            Rp02napr /= 4;

            // шпилька стяжная
            double Rp02shpilka;
            double.TryParse(report.GetValue("Table2_PredelTech_52"), out Rp02shpilka);
            Rp02shpilka /= 4;

            // 1
            report.SetValue("Table13_TauSquareP_1", DigitalProcess.MathTrancate(Rp02bolt, 1, false));
            TauSquareP[0] = Rp02bolt;

            // 2
            report.SetValue("Table13_TauSquareP_2", DigitalProcess.MathTrancate(Rp02plita, 1, false));
            TauSquareP[1] = Rp02plita;

            // 3
            report.SetValue("Table13_TauSquareP_3", DigitalProcess.MathTrancate(Rp02bolt, 1, false));
            TauSquareP[2] = Rp02bolt;

            // 4
            report.SetValue("Table13_TauSquareP_4", DigitalProcess.MathTrancate(Rp02napr, 1, false));
            TauSquareP[3] = Rp02napr;

            // 5
            report.SetValue("Table13_TauSquareP_5", DigitalProcess.MathTrancate(Rp02bolt, 1, false));
            TauSquareP[4] = Rp02bolt;

            // 6
            report.SetValue("Table13_TauSquareP_6", DigitalProcess.MathTrancate(Rp02napr, 1, false));
            TauSquareP[5] = Rp02napr;

            // 7
            report.SetValue("Table13_TauSquareP_7", DigitalProcess.MathTrancate(Rp02shpilka, 1, false));
            TauSquareP[6] = Rp02shpilka;

            // 8
            report.SetValue("Table13_TauSquareP_8", DigitalProcess.MathTrancate(Rp02gaika, 1, false));
            TauSquareP[7] = Rp02gaika;

            bool bWithoutFlange = (comboBox16.SelectedIndex > 2);
            const string c_space = "-";

            if (bWithoutFlange)
            {
                // F
                report.SetValue("Table13_F_1", c_space);
                report.SetValue("Table13_F_2", c_space);

                // d
                report.SetValue("Table13_d_1", c_space);
                report.SetValue("Table13_d_2", c_space);

                // h
                report.SetValue("Table13_h_1", c_space);
                report.SetValue("Table13_h_2", c_space);

                // k1
                report.SetValue("Table13_K1_1", c_space);
                report.SetValue("Table13_K1_2", c_space);

                // z
                report.SetValue("Table13_z_1", c_space);
                report.SetValue("Table13_z_2", c_space);

                // tau
                report.SetValue("Table13_TauP_1", c_space);
                report.SetValue("Table13_TauP_2", c_space);

                // tau square
                report.SetValue("Table13_TauSquareP_1", c_space);
                report.SetValue("Table13_TauSquareP_2", c_space);
            }

            bool Condition = true;

            // условие прочности
            for (i = 0; i < 8; i++)
            {
                if (TauP[i] > TauSquareP[i])
                {
                    Condition = false;
                    break;
                }
            }

            if (!Condition)
            {
                m_messages.Add("Таблица 13, стр. 27");
            }

            return true;
        }

        private bool FillTable14(ref SReport_Utility.PrintReport report)
        {
            string ptoName = comboBox1.Text;

            // P
            double P;
            double.TryParse(textBox1.Text, out P);

            report.SetValue("Table14_P", P.ToString());

            // F
            double F;
            double.TryParse(report.GetValue("Table10_Fwp_1"), out F);
            report.SetValue("Table14_F", F.ToString());

            // Apr
            double Am = m_hashTablePlit.GetValue(ptoName, "Am");
            report.SetValue("Table14_Apr", Am.ToString());

            // q
            double q = P * Am;
            report.SetValue("Table14_q", q.ToString());

            // L
            double L = m_hashTablePlit.GetValue(ptoName, "Bsh");
            report.SetValue("Table14_L", L.ToString());

            // z
            double z1 = m_list55.GetValue(ptoName, "z1");
            double z2 = m_list55.GetValue(ptoName, "z2");
            double z3 = m_list55.GetValue(ptoName, "z3");

            report.SetValue("Table14_z1", z1.ToString());
            report.SetValue("Table14_z2", z2.ToString());
            report.SetValue("Table14_z3", z3.ToString());

            // Bpr
            double Bm = m_hashTablePlit.GetValue(ptoName, "Bm");
            report.SetValue("Table14_Bpr", Bm.ToString());

            // a
            double a = (L - Bm) / 2;
            report.SetValue("Table14_a", a.ToString());

            // M
            double M1, M2, M3;

            M1 = F * a + q * (z1 * (L - z1) - a * a) / 2;
            report.SetValue("Table14_M1", DigitalProcess.MathTrancate(Convert.ToInt32(M1), 2).ToString());

            M2 = F * a + q * (z2 * (L - z2) - a * a) / 2;
            report.SetValue("Table14_M2", DigitalProcess.MathTrancate(Convert.ToInt32(M2), 2).ToString());

            M3 = F * a + q * (z3 * (L - z3) - a * a) / 2;
            report.SetValue("Table14_M3", DigitalProcess.MathTrancate(Convert.ToInt32(M3), 2).ToString());

            // b
            double b1 = m_list55.GetValue(ptoName, "b1");
            double b2 = m_list55.GetValue(ptoName, "b2");
            double b3 = m_list55.GetValue(ptoName, "b3");

            report.SetValue("Table14_b1", b1.ToString());
            report.SetValue("Table14_b2", b2.ToString());
            report.SetValue("Table14_b3", b3.ToString());

            // h
            double h = m_hashTablePlit.GetValue(ptoName, "S1");
            report.SetValue("Table14_h", h.ToString());

            // W
            double W1, W2, W3;

            W1 = b1 * h * h / 6;
            report.SetValue("Table14_W1", GetDoubleFormat(W1, 0));

            W2 = b2 * h * h / 6;
            report.SetValue("Table14_W2", GetDoubleFormat(W2, 0));

            W3 = b3 * h * h / 6;
            report.SetValue("Table14_W3", GetDoubleFormat(W3, 0));

            // Sigma square
            double SigmaSquare;
            double.TryParse(report.GetValue("Table2_Sigma_12"), out SigmaSquare);

            report.SetValue("Table14_SigmaSquare", SigmaSquare.ToString());

            // sigma isgib
            double sigma1 = M1 / W1;
            double sigma2 = M2 / W2;
            double sigma3 = M3 / W3;

            report.SetValue("Table14_SigmaIs1", DigitalProcess.MathTrancate(sigma1, 1, false));
            report.SetValue("Table14_SigmaIs2", DigitalProcess.MathTrancate(sigma2, 1, false));
            report.SetValue("Table14_SigmaIs3", DigitalProcess.MathTrancate(sigma3, 1, false));

            // Sigma square isgib
            SigmaSquare *= 1.3;
            report.SetValue("Table14_SSIs", SigmaSquare.ToString());

            // условие прочности
            bool Condition = (SigmaSquare > sigma1);
            Condition = Condition && (SigmaSquare > sigma2);
            Condition = Condition && (SigmaSquare > sigma3);

            if (!Condition)
            {
                m_messages.Add("Таблица 14, стр. 30");
            }

            // Рисунок 7
            double relation = 9.53 / 14.72;
            Bitmap bitmap = SReport_Utility.SimpleImageProccessing.Zoom(Properties.Resources.M_Universal, 1816, Convert.ToInt32(1816 * relation));

            report.SetPicture("Picture7", bitmap);

            return true;
        }

        private bool FillTable15(ref SReport_Utility.PrintReport report)
        {
            string ptoName = comboBox1.Text;

            // A1
            double A1 = m_list56.GetValue(ptoName, "A1");
            report.SetValue("Table15_A_1", A1.ToString());

            // A2
            double A2 = m_list56.GetValue(ptoName, "A2");
            report.SetValue("Table15_A_2", A2.ToString());

            // Fw
            double Fw;
            double.TryParse(report.GetValue("Table10_Fwp_4"), out Fw);

            // z
            double z;
            double.TryParse(report.GetValue("Table5_z"), out z);

            // Sigma 1
            double Sigma1 = Fw / (z * A1);
            report.SetValue("Table15_Sigma_1", DigitalProcess.MathTrancate(Sigma1, 1, false));

            // Sigma 2
            double Sigma2 = Fw / (z * A2);
            report.SetValue("Table15_Sigma_2", DigitalProcess.MathTrancate(Sigma2, 1, false));

            // Sigma Square
            double sigma;
            double.TryParse(report.GetValue("Table2_Sigma_12"), out sigma);
            sigma *= 2;
            report.SetValue("Table15_SigmaSquare", DigitalProcess.MathTrancate(sigma, 1, false));

            // условие прочности
            bool Condition = (sigma > Sigma1);
            Condition = Condition && (sigma > Sigma2);

            if (!Condition)
            {
                m_messages.Add("Таблица 15, стр. 31");
            }

            return true;
        }

        private bool FillTable16(ref SReport_Utility.PrintReport report)
        {
            string ptoName = comboBox1.Text;

            // "N1", "N2", "N3", "NSigma", "NN", "A", "B", "K"

            // N sigma
            double Nsigma = m_list57.GetValue(ptoName, "NSigma");
            report.SetValue("Table16_NSigma", Nsigma.ToString());

            // nN
            double nN = m_list57.GetValue(ptoName, "NN");
            report.SetValue("Table16_NN", nN.ToString());

            // A
            double A = m_list57.GetValue(ptoName, "A");
            report.SetValue("Table16_A", A.ToString());

            // B
            double B, Rp02;            
            double.TryParse(report.GetValue("Table2_PredelTech_52"), out Rp02);
            B = m_list57.GetValue(ptoName, "B");
            B *= Rp02;
            report.SetValue("Table16_B", DigitalProcess.MathTrancate(B, 2, false));

            // t
            double t;
            double.TryParse(textBox11.Text, out t);
            report.SetValue("Table16_t", t.ToString());

            // K
            double K = m_list57.GetValue(ptoName, "K");
            report.SetValue("Table16_K", K.ToString());

            // Sigma 4w
            double Sigma4w1, Sigma4w2, Sigma4w3;

            double.TryParse(report.GetValue("Table10_Sigma4W_Round_1"), out Sigma4w1);
            double.TryParse(report.GetValue("Table10_Sigma4W_Round_3"), out Sigma4w2);
            double.TryParse(report.GetValue("Table10_Sigma4W_Round_4"), out Sigma4w3);

            report.SetValue("Table16_Sigma4w_1", GetDoubleFormat(Sigma4w1, 0));
            report.SetValue("Table16_Sigma4w_2", GetDoubleFormat(Sigma4w2, 0));
            report.SetValue("Table16_Sigma4w_3", GetDoubleFormat(Sigma4w3, 0));

            // Sigma A
            double SigmaA1 = K * Sigma4w1;
            double SigmaA2 = K * Sigma4w2;
            double SigmaA3 = K * Sigma4w3;

            report.SetValue("Table16_SigmaA_1", GetDoubleFormat(SigmaA1, 0));
            report.SetValue("Table16_SigmaA_2", GetDoubleFormat(SigmaA2, 0));
            report.SetValue("Table16_SigmaA_3", GetDoubleFormat(SigmaA3, 0));

            // N square
            double s;
            double Nsquare1, Nsquare2, Nsquare3;

            // 1
            s = A * (2300 - t) / ((SigmaA1 - B / Nsigma) * 2300);
            Nsquare1 = s * s / nN;

            report.SetValue("Table16_NSquare_1", GetDoubleFormat(Nsquare1, 0));

            // 2
            s = A * (2300 - t) / ((SigmaA2 - B / Nsigma) * 2300);
            Nsquare2 = s * s / nN;

            report.SetValue("Table16_NSquare_2", GetDoubleFormat(Nsquare2, 0));

            // 3
            s = A * (2300 - t) / ((SigmaA3 - B / Nsigma) * 2300);
            Nsquare3 = s * s / nN;

            report.SetValue("Table16_NSquare_3", GetDoubleFormat(Nsquare3, 0));

            // N
            double N2 = m_list57.GetValue(ptoName, "N2");
            double N1 = m_list57.GetValue(ptoName, "N1") * N2;
            double N3 = m_list57.GetValue(ptoName, "N3");

            report.SetValue("Table16_N1", N1.ToString());
            report.SetValue("Table16_N2", N2.ToString());
            report.SetValue("Table16_N3", N3.ToString());

            // a
            double al = N1 / Nsquare1 + N2 / Nsquare2 + N3 / Nsquare3;
            report.SetValue("Table16_al", DigitalProcess.MathTrancate(al, 3, false));

            // N1m
            report.SetValue("Table16_N1M", m_list57.GetValue(ptoName, "N1").ToString());

            // Bm
            report.SetValue("Table16_BM", m_list57.GetValue(ptoName, "B").ToString());

            // условие прочности
            if (al > 1)
            {
                m_messages.Add("Таблица 16, стр. 33");
            }

            return true;
        }

        private bool IsValid()
        {
            m_lastError = "";
            bool result = false;

            do
            {
                int nVal;

                // количество пластин
                if (!int.TryParse(textBox3.Text, out nVal))
                {
                    m_lastError = "Количество пластин для ПТО задано не корректно.";
                    break;
                }

                // заводской номер
                if (textBox4.Text.Length < 5)
                {
                    m_lastError = "Заводской номер ПТО задан не корректно.";
                    break;
                }

                double dVal;
                double T;

                // расчетная температура
                if (!double.TryParse(textBox11.Text, out T))
                {
                    m_lastError = "Расчетная температрута задана не корректно.";
                    break;
                }

                // расчетное давление
                if (!double.TryParse(textBox1.Text, out dVal))
                {
                    m_lastError = "Расчетное давление задано не корректно.";
                    break;
                }

                // давление гидроиспытаний
                if (!double.TryParse(textBox2.Text, out dVal))
                {
                    m_lastError = "Давлние гидроиспытаний задано не корректно.";
                    break;
                }

                // масса пустой
                double MNetto;
                if (!double.TryParse(textBox10.Text, out MNetto))
                {
                    m_lastError = "Масса пустого ПТО задана не корректно.";
                    break;
                }

                // масса заполненный
                double MBrutto;
                if (!double.TryParse(textBox9.Text, out MBrutto))
                {
                    m_lastError = "Масса заполненного ПТО задана не корректно.";
                    break;
                }

                // сравниваем брутто и нетто
                if (MNetto >= MBrutto)
                {
                    m_lastError = "Масса заполненного ПТО должна превышать массу пустого.";
                    break;
                }

                double MaxT;

                // максимально-допустимая температура, горячий контур
                if (!double.TryParse(textBox7.Text, out MaxT))
                {
                    m_lastError = "Максимально допустимая температура по горячему контуру задана не корректно.";
                    break;
                }

                if (MaxT > T)
                {
                    m_lastError = "Максимально допустимая температура по горячему контуру превосходит расчетную температуру.";
                    break;
                }

                // максимально-допустимая температура, холодный контур
                if (!double.TryParse(textBox12.Text, out MaxT))
                {
                    m_lastError = "Максимально допустимая температура по холодному контуру задана не корректно.";
                    break;
                }

                if (MaxT > T)
                {
                    m_lastError = "Максимально допустимая температура по холодному контуру превосходит расчетную температуру.";
                    break;
                }

                // минимально-допустимая температура
                if (!double.TryParse(textBox8.Text, out dVal))
                {
                    m_lastError = "Минимально допустимая температура задана не корректно.";
                    break;
                }

                // Объем ПТО
                if (!double.TryParse(textBox5.Text, out dVal))
                {
                    m_lastError = "Объем ПТО задан не корректно.";
                    break;
                }
               
                result = true;

            } while (false);

            return result;
        }

        private bool ReadExcelData()
        {
            string sData = Application.StartupPath + "\\data.xls";
            //string SharedPath = "X:\\Павлов\\Расчет на прочность\\Data\\data.xls";

            //if (File.Exists(SharedPath))
            //{
            //    File.Copy(SharedPath, sData, true);
            //}

            if (File.Exists(sData))
            {
                // читаем данные из Excel

                m_hashTable43 = ReadHashTable43(sData);
                m_hashTable423 = ReadHashTable423(sData);
                m_hashTable422 = ReadHashTable422(sData);
                m_hashTablePlit = ReadHashTable41(sData);
                m_hashTable424 = ReadHashTable424(sData);
                m_hashTable51 = ReadHashTable51(sData);
                m_list52 = ReadList52(sData);
                m_len52 = ReadLen52(sData);
                m_list53 = ReadList53(sData);
                m_list54 = ReadList54(sData);
                m_list55 = ReadList55(sData);
                m_list56 = ReadList56(sData);
                m_list57 = ReadList57(sData);

                m_T = new double[6];
                ReadTemperature(sData);
            }
            else
            {
                MessageBox.Show("Файл данных не обнаружен.", KitConstant.companyName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Close();
            }

            return true;
        }

        private void FillHashTable(ref double[,] DoubleValue, ref MyDoubleHashTable hashTable, int len)
        {
            double[] data = new double[len];

            for (int j = 0; j < c_pto_count; j++)
            {
                for (int i = 0; i < len; i++)
                    data[i] = DoubleValue[j, i];

                hashTable.AddValue(m_pto_list[j], data);
            }
        }

        private MyDoubleHashTable ReadHashTable43(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = new string[] { "f", "n", "Am", "Bm", "A1", "B1", "b", "h", "dh", "epsi", "Epr", "q", "R", "ps" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list43";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadHashTable41(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "Bm", "B", "Am", "A", "b1", "Bsh", "d", "z", "c111", "c112", "c12", "c2", "S1", "S2" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list41";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadHashTable422(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "q0liq", "q0vap", "q0air", "mliq", "mvap", "mair", "eta", "hi", "z", "d0", "d0w", "ksi" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list422";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadHashTable423(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "d0up", "d0wup", "Mkup", "zc", "d0down", "d0wdown", "Мкdown", "zdown", "ksi" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list423";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadHashTable51(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "k4", "k5", "k6", "k7", "k8", "k9" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list51";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadHashTable424(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "b01", "Dm1", "b02", "Dm2", "b03", "Dm3", "b1rib", "b2rib", "b3rib", "deltarib", "q0liqrib", "q0vaprib", "q0airrib", "mliqrib", "mvaprib", "mairrib", "nurib", "b1paronit", "b2paronit", "b3paronit", "deltaparonit", "q0liqparonit", "q0vapparonit", "q0airparonit", "mliqparonit", "mvapparonit", "mairparonit", "nuparonit", "hi", "ksi", "dc", "z", "d0", "d0w" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list424";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadList52(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "aisim4", "aisim5",	"aisim6", "aisim7", "aisim8", "aisim9", "titanm5", "titanm6", "titanm7", "titanm8", "titanm9", "hasm5", "hasm6", "hasm7", "hasm8", "hasm9", "epdm", "nitril", "viton", "F" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list52";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadLen52(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "W1250", "W1330", "W1400", "W1500", "W1600", "W1800", "W1900", "W11000", "W11200", "W11300", "W11500", "W12000", "W12500", "W13000", "W14000", "W15000", "W16000", "W2", "W1750" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "len52";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            //double val = hashTable.GetValue(comboBox1.Text, "W1600");

            return hashTable;
        }

        private MyDoubleHashTable ReadList54(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "K1sh", "K1g", "Km", "h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list54";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            //double val = hashTable.GetValue(comboBox1.Text, "K1sh");

            return hashTable;
        }

        private MyDoubleHashTable ReadList55(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "z1", "z2", "z3", "b1", "b2", "b3" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list55";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadList56(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "A1", "A2" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list56";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadList57(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "N1", "N2", "N3", "NSigma", "NN", "A", "B", "K" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list57";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private MyDoubleHashTable ReadList53(string source)
        {
            // declare excel reader
            ExcelReader excelReader = new ExcelReader();
            string[] fields = { "mu" };

            excelReader.DbPath = source;
            excelReader.Fields = fields;
            excelReader.Page = "list53";

            double[,] DoubleValue = excelReader.GetDoubleTable();

            // === setup data ===
            MyDoubleHashTable hashTable = new MyDoubleHashTable(fields);
            FillHashTable(ref DoubleValue, ref hashTable, fields.Length);

            return hashTable;
        }

        private void ReadTemperature(string source)
        {
            // declare excel reader
            string[] fields = { "t1", "t2", "t3", "t4", "t5", "t6" };
            ExcelReader excelReader = new ExcelReader(source, fields, "Temperature");

            double[,] DoubleValue = excelReader.GetDoubleTable();
            int len = m_T.Length;

            // === setup data ===
            for (int i = 0; i < len; i++)
            {
                m_T[i] = DoubleValue[0, i];
            }

            ReadSteelProperties(ref DoubleValue, len);
        }

        private MyIntegerHashTable CreateLen()
        {
            int[] data = null;
            int index = 0, i;
            MyIntegerHashTable hashTable = new MyIntegerHashTable();

            // 04
            data = new int[] { 250, 330, 500 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 08
            //data = new int[] { 250, 330, 500 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 07
            data = new int[] { 400, 600, 800 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 14
            //data = new int[] { 400, 600, 800 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 20
            //data = new int[] { 400, 600, 800 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 21
            data = new int[] { 500, 800, 1200, 1500 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 22
            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 47
            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 41
            data = new int[] { 600, 1000, 1500, 2000, 2500, 3000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 42
            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 62
            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 86
            data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 110
            //data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 43
            data = new int[] { 600, 1000, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 65
            //data = new int[] { 600, 1000, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 100            
            data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 130
            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 152
            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 220
            data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 113
            //data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 81
            data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 121
            //data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 188
            //data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 251
            //data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 145
            data = new int[] { 1000, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 4; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 210
            //data = new int[] { 1000, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 4; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 201
            //data = new int[] { 1000, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 2; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 19
            data = new int[] { 400, 500, 600, 750, 1000 };

            for (i = 0; i < 4; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 53
            data = new int[] { 1000, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 2; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 160
            data = new int[] { 1000, 1500, 2000, 2500, 3000, 4000, 5000, 6000 };

            for (i = 0; i < 2; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 19
            data = new int[] { 400, 500, 600, 750, 800, 1000 };

            for (i = 0; i < 8; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 26
            data = new int[] { 600, 1000, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 6; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // XG-31
            data = new int[] { 400, 600, 750, 900 };

            for (i = 0; i < 3; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            // 229
            data = new int[] { 600, 1000, 1300, 1500, 2000, 2500, 3000, 4000 };

            for (i = 0; i < 2; i++)
            {
                hashTable.AddValue(m_pto_list[index++], data);
            }

            return hashTable;
        }

        private void ReadSteelProperties(ref double[,] DoubleValue, int len)
        {
            int i;
            int index = 1;
            double[] Rp = new double[len];
            double[] Rm = new double[len];
            double[] E = new double[len];
            double[] Alpha = new double[len];

            // plits
            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_PlitSt3 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Plit09G2C = new SteelProperty(Rp, Rm, E, Alpha);

            // material - napr-up
            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprUpSt2 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprUpSt3 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprUp20 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprUp20x13 = new SteelProperty(Rp, Rm, E, Alpha);

            // material - napr-down
            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprDownSt2 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprDownSt3 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprDown20 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_NaprDown20x13 = new SteelProperty(Rp, Rm, E, Alpha);

            // material - krepeg
            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Krepeg40x = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Krepeg35 = new SteelProperty(Rp, Rm, E, Alpha);

            // material - rezba
            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Rezba09G2C = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Rezba20x13 = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Rezba20 = new SteelProperty(Rp, Rm, E, Alpha);

            // plates
            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_AISI = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Titan = new SteelProperty(Rp, Rm, E, Alpha);

            for (i = 0; i < len; i++) Rp[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Rm[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) E[i] = DoubleValue[index, i]; index++;
            for (i = 0; i < len; i++) Alpha[i] = DoubleValue[index, i]; index++;
            index++;

            m_Steel45 = new SteelProperty(Rp, Rm, E, Alpha);
        }

        #endregion

        private void OnKeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back || e.KeyChar == m_nfi.NumberDecimalSeparator[0]))
            {
                e.Handled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            m_nfi = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat;

            foreach (PrimaryFeature ph in m_manager)
            {
                contextMenuStrip1.Items.Add(ph.Name);
                contextMenuStrip2.Items.Add(ph.Name);
            }

            StartSets();
            CreatePTOList();
            m_ptoHashLen = CreateLen();

            // === PTO ===
            comboBox1.BeginUpdate();
            comboBox1.Items.AddRange(m_pto_list.ToArray());
            comboBox1.SelectedIndex = 31;
            comboBox1.EndUpdate();

            comboBox1.Select();
            // ===========

            // Read the data from Excel
            ReadExcelData();

            // === Материал пластин ===
            comboBox2.Items.Add(Utility.KitConstant.AISI_316);
            comboBox2.Items.Add(Utility.KitConstant.SMO_254);
            comboBox2.Items.Add(Utility.KitConstant.Titan);
            comboBox2.SelectedIndex = 0;
            // ========================

            // === Толщина пластин ===
            foreach (double s in m_s)
            {
                comboBox3.Items.Add(string.Format("{0} mm", s));
            }

            comboBox3.SelectedIndex = 1;
            // =======================

            // === Фланцы ===
            comboBox16.Items.Add("1");
            comboBox16.Items.Add("2");
            comboBox16.Items.Add("4");
            comboBox16.Items.Add("Без ответных фланцев");
            comboBox16.SelectedIndex = 0;
            // ==============

            // === Материал прокладок ===
            Utility.WrapperManager wrapMan = new Utility.WrapperManager();

            comboBoxWrapper.BeginUpdate();
            comboBoxWrapper.Items.AddRange(wrapMan.items);
            comboBoxWrapper.SelectedIndex = 0;
            comboBoxWrapper.EndUpdate();
            // ==========================

            // === Материалы и комплектующие ===
            comboBox4.Items.Add("ПОН-Б");
            comboBox4.Items.Add("ПМБ");
            comboBox4.Items.Add("ИРП");
            comboBox4.SelectedIndex = 0;

            comboBox5.Items.Add(KitConstant.Steel_st3);
            comboBox5.Items.Add(KitConstant.Steel_09G2C);
            comboBox5.SelectedIndex = 0;

            comboBox6.Items.Add(KitConstant.Steel_st3);
            comboBox6.Items.Add(KitConstant.Steel_09G2C);
            comboBox6.SelectedIndex = 0;

            m_HashSteel.Add(KitConstant.Steel_st2, "ГОСТ 380-2005");
            m_HashSteel.Add(KitConstant.Steel_st3, "ГОСТ 380-2005");
            m_HashSteel.Add(KitConstant.Steel_20, "ГОСТ 1050-2013");
            m_HashSteel.Add(KitConstant.Steel_20x13, "ГОСТ 5949-75");
            m_HashSteel.Add(KitConstant.Steel_40x, "ГОСТ 4543-71", "КП 8");
            m_HashSteel.Add(KitConstant.Steel_35, "ГОСТ 1050-2013", "КП 8");
            m_HashSteel.Add(KitConstant.Steel_09G2C, "ГОСТ 19281-2014");
            m_HashSteel.Add(KitConstant.Steel_45, "ГОСТ 1050-2013", "КП 8");
            
            comboBox7.Items.Add(KitConstant.Steel_st2);
            comboBox7.Items.Add(KitConstant.Steel_st3);
            comboBox7.Items.Add(KitConstant.Steel_20);
            comboBox7.Items.Add(KitConstant.Steel_20x13);
            comboBox7.Items.Add(KitConstant.Steel_45);
            comboBox7.SelectedIndex = 0;

            comboBox8.Items.Add(KitConstant.Steel_st2);
            comboBox8.Items.Add(KitConstant.Steel_st3);
            comboBox8.Items.Add(KitConstant.Steel_20);
            comboBox8.Items.Add(KitConstant.Steel_20x13);
            comboBox8.Items.Add(KitConstant.Steel_45);
            comboBox8.SelectedIndex = 0;

            comboBox9.Items.Add(KitConstant.Steel_40x);
            comboBox9.SelectedIndex = 0;

            comboBox10.Items.Add(KitConstant.Steel_35);
            comboBox10.Items.Add(KitConstant.Steel_45);
            comboBox10.Items.Add(KitConstant.Steel_40x);
            comboBox10.SelectedIndex = 0;

            comboBox11.Items.Add(KitConstant.Steel_40x);
            comboBox11.SelectedIndex = 0;

            comboBox12.Items.Add(KitConstant.Steel_09G2C);
            comboBox12.Items.Add(KitConstant.Steel_20x13);
            comboBox12.Items.Add(KitConstant.Steel_20);
            comboBox12.Items.Add(KitConstant.Steel_45);
            comboBox12.SelectedIndex = 0;

            comboBox13.SelectedIndex = 0;
            comboBox14.SelectedIndex = 0;
            // ================================= 

            // === fill text boxes ===
            textBox3.Text = "22";

            // === Ответственные ===
            string sConfig = Source_Path + "\\people.txt";

            if (File.Exists(sConfig))
            {
                StreamReader sr = new StreamReader(sConfig);

                textBox13.Text = sr.ReadLine();
                textBox14.Text = sr.ReadLine();
                textBox15.Text = sr.ReadLine();
                textBox16.Text = sr.ReadLine();
                textBox17.Text = sr.ReadLine();

                sr.Close();
            }
            // =====================

            textBox4.Text = "12345"; // короткий заводской номер
            textBox11.Text = "60"; // расчетная температура
            textBox1.Text = "1,6"; // расчетное давление
            textBox2.Text = "2,3"; // давление гидроиспытаний

            textBox10.Text = "25"; // масса пустой
            textBox9.Text = "30"; // масса заполненный

            textBox7.Text = "50"; // максимально-допустимая температура, горячий контур
            textBox12.Text = "50"; // максимально-допустимая температура, холодный контур
            textBox8.Text = "-20"; // минимально-допустимая температура

            textBox5.Text = "1"; // объем ПТО

            // fill dictionaries
            FillDictSteel();

            // === title ===
            this.Text = string.Format("{0}: Новый", KitConstant.softwareName);
            lblVersion.Text = KitConstant.companyName + ", версия: " + VersionControl_VB.MyVersionControl.MyVersion;
            // ===============
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!IsValid())
            {
                MessageBox.Show(m_lastError, string.Format("{0}: {1}", KitConstant.companyName, KitConstant.softwareName), MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            // === Original code ===
            Stimulsoft.Report.StiReport stireport = new Stimulsoft.Report.StiReport();
            PrintReport report = new PrintReport();

            m_messages.Clear();

            if (!FillTables(ref report))
            {
                MessageBox.Show(m_lastError, KitConstant.softwareName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (m_messages.Count > 0)
            {
                Form2 dlg = new Form2(m_messages);

                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
            }
            // ==================

            stireport.RegData(report.Table);

            int index = comboBox15.SelectedIndex;
           
            string sPath = Application.StartupPath + "\\reports\\"; //myreport_" + index.ToString();

            if (IsBig())
            {
                sPath += "big\\myreport_" + index.ToString();
                sPath += "_new_big.mrt";
            }
            else if (Is0408())
            {
                sPath += "0408\\myreport_" + index.ToString();
                sPath += "_new_0408.mrt";
            }
            else if (Is212247())
            {
                sPath += "212247\\myreport_" + index.ToString();
                sPath += "_212247.mrt";
            }
            else
            {
                sPath += "new\\myreport_" + index.ToString();
                sPath += "_new.mrt";
            }

            stireport.Load(sPath);

            if (Form.ModifierKeys == Keys.Shift)
            {
                stireport.Design();
            }
            else
            {
                Stimulsoft.Report.Components.StiPagesCollection pageCol = stireport.Pages;
                System.Collections.IEnumerator pageEn = pageCol.GetEnumerator();

                while (pageEn.MoveNext())
                {
                    Stimulsoft.Report.Components.StiPage page = pageEn.Current as Stimulsoft.Report.Components.StiPage;
                    Stimulsoft.Report.Components.StiComponentsCollection cc = page.Components;

                    System.Collections.IEnumerator compEn = cc.GetEnumerator();

                    while (compEn.MoveNext())
                    {
                        Stimulsoft.Report.Components.StiComponent component = compEn.Current as Stimulsoft.Report.Components.StiComponent;

                        if (component is Stimulsoft.Report.Components.IStiText)
                        {
                            Stimulsoft.Report.Components.IStiText text = component as Stimulsoft.Report.Components.IStiText;
                            string value = text.Text;

                            if (!value.EndsWith(" "))
                            {
                                value += " ";
                                text.Text = value;
                            }
                        }
                    }

                    page.Components = cc;
                }

                stireport.Show();
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back))
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back))
            {
                e.Handled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ptoName = comboBox1.Text;
            int[] len = m_ptoHashLen.GetValue(ptoName);

            comboBox17.BeginUpdate();
            comboBox17.Items.Clear();

            foreach (int x in len)
            {
                comboBox17.Items.Add(x);
            }

            comboBox17.SelectedIndex = 0;
            comboBox17.EndUpdate();

            // наполнение по страницам
            comboBox15.BeginUpdate();
            comboBox15.Items.Clear();

            if (!Is0408())
            {
                comboBox15.Items.Add("Этикетка"); // 0
                comboBox15.Items.Add("1-4"); // 1
                comboBox15.Items.Add("5-6"); // 2
                comboBox15.Items.Add("7-8"); // 3
                comboBox15.Items.Add("9-10"); // 4
                comboBox15.Items.Add("11-12"); // 5
                comboBox15.Items.Add("13-14"); // 6
                comboBox15.Items.Add("15-16"); // 7
                comboBox15.Items.Add("17-18"); // 8
                comboBox15.Items.Add("19-20"); // 9
                comboBox15.Items.Add("21-22"); // 10
                comboBox15.Items.Add("23-24"); // 11
                comboBox15.Items.Add("25-26"); // 12
                comboBox15.Items.Add("27-30"); // 13
                comboBox15.Items.Add("31-32"); // 14
                comboBox15.Items.Add("33-36"); // 15
                //comboBox15.Items.Add("Frame"); // 17

                // тип фланцевого соединения
                comboBox16.Enabled = true;
            }
            else
            {
                // 04-08
                comboBox15.Items.Add("Этикетка"); // 0
                comboBox15.Items.Add("1-4"); // 1
                comboBox15.Items.Add("5-6"); // 2
                comboBox15.Items.Add("7-8"); // 3
                comboBox15.Items.Add("9-10"); // 4
                comboBox15.Items.Add("11-12"); // 5
                comboBox15.Items.Add("13-14"); // 6
                comboBox15.Items.Add("15-16"); // 7
                comboBox15.Items.Add("17-18"); // 8
                comboBox15.Items.Add("19-20"); // 9
                comboBox15.Items.Add("21-22"); // 10
                comboBox15.Items.Add("23-24"); // 11
                comboBox15.Items.Add("25-28"); // 12
                comboBox15.Items.Add("29-30"); // 13
                comboBox15.Items.Add("31-34"); // 14

                // тип фланцевого соединения
                comboBox16.Enabled = false;
            }

            comboBox15.SelectedIndex = 0;
            comboBox15.EndUpdate();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!IsValid())
            {
                MessageBox.Show(m_lastError, string.Format("{0}: {1}", KitConstant.companyName, KitConstant.softwareName), MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            PrintReport report = new PrintReport();
            m_messages.Clear();

            if (!FillTables(ref report))
            {
                MessageBox.Show(m_lastError, SReport_Utility.KitConstant.softwareName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (m_messages.Count > 0)
            {
                Form2 dlg = new Form2(m_messages);

                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
            }
            // ==================

            PrintDialog prnDlg = new PrintDialog();
            prnDlg.AllowPrintToFile = false;

            if (prnDlg.ShowDialog() == DialogResult.OK)
            {
                // create file collection
                System.Collections.ArrayList reportList = new System.Collections.ArrayList();
                
                string sExt;
                int count;
                string startUp = Application.StartupPath + "\\reports\\";

                if (IsBig())
                {
                    startUp += "big\\";
                    sExt = "_new_big.mrt";
                    count = 16;
                }
                else if (Is0408())
                {
                    startUp += "0408\\";
                    sExt = "_new_0408.mrt";
                    count = 15;
                }
                else if (Is212247())
                {
                    startUp += "212247\\";
                    sExt = "_212247.mrt";
                    count = 16;
                }
                else
                {
                    startUp += "new\\";
                    sExt = "_new.mrt";
                    count = 16;
                }

                // fill file collection
                for (int i = 0; i < count; i++)
                {
                    reportList.Add(startUp + "myreport_" + i.ToString() + sExt);
                }

                foreach (string file in reportList)
                {
                    Stimulsoft.Report.StiReport stireport = new Stimulsoft.Report.StiReport();

                    stireport.Load(file);
                    stireport.RegData(report.Table);

                    Stimulsoft.Report.Components.StiPagesCollection pageCol = stireport.Pages;
                    System.Collections.IEnumerator pageEn = pageCol.GetEnumerator();

                    while (pageEn.MoveNext())
                    {
                        Stimulsoft.Report.Components.StiPage page = pageEn.Current as Stimulsoft.Report.Components.StiPage;
                        Stimulsoft.Report.Components.StiComponentsCollection cc = page.Components;

                        System.Collections.IEnumerator compEn = cc.GetEnumerator();

                        while (compEn.MoveNext())
                        {
                            Stimulsoft.Report.Components.StiComponent component = compEn.Current as Stimulsoft.Report.Components.StiComponent;

                            if (component is Stimulsoft.Report.Components.IStiText)
                            {
                                Stimulsoft.Report.Components.IStiText text = component as Stimulsoft.Report.Components.IStiText;
                                string value = text.Text;

                                if (!value.EndsWith(" "))
                                {
                                    value += " ";
                                    text.Text = value;
                                }
                            }
                        }

                        page.Components = cc;
                    }

                    stireport.Print(false, prnDlg.PrinterSettings);
                    //stireport.Clear();

                    System.Threading.Thread.Sleep(250);
                }
            }
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ==========================
            if (IsBig())
            {                
                Set set = new Set();

                set.AddRange(48, 65); // 41-42, 62
                set.AddRange(84, 89); // 65
                set.AddRange(66, 71); // 86
                set.AddRange(90, 95); // 100
                set.AddRange(72, 77); // 110
                set.AddRange(114, 119); // 113
                set.AddRange(102, 107); // 152

                int L;
                int.TryParse(comboBox17.Text, out L);

                if (L <= 2000)
                {
                    set.AddRange(152, 153); // 201
                }

                if (L <= 2800)
                {
                    set.AddRange(148, 151); // 210
                }

                if (L <= 2000)
                {
                    set.AddRange(108, 113); // 220
                }

                int index = comboBox1.SelectedIndex;
                m_bNaprIn = set.In(index);
            }
            else
            {
                m_bNaprIn = true;
            }
            // ==============================
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            string startupPath = Source_Path + "\\";

            // === Ответственные ===
            using (StreamWriter sw2 = new StreamWriter(startupPath + "people.txt"))
            {
                sw2.WriteLine(textBox13.Text);
                sw2.WriteLine(textBox14.Text);
                sw2.WriteLine(textBox15.Text);
                sw2.WriteLine(textBox16.Text);
                sw2.WriteLine(textBox17.Text);
            }
            // =====================
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmSave dlg = new frmSave();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                // save
                SReportArchive archive = new SReportArchive();

                archive.SetValue("PTO_Type", comboBox1.SelectedIndex);
                archive.SetValue("PTO_PlateCount", textBox3.Text);
                archive.SetValue("PTO_PlateMaterial", comboBox2.SelectedIndex);
                archive.SetValue("PTO_PlateThickness", comboBox3.SelectedIndex);
                archive.SetValue("PTO_WrapperMaterial1", comboBoxWrapper.SelectedIndex);
                archive.SetValue("PTO_WrapperMaterial2", comboBox4.SelectedIndex);

                archive.SetValue("PTO_Volume", textBox5.Text);
                archive.SetValue("PTO_Len", comboBox17.SelectedIndex);
                archive.SetValue("PTO_WeightNetto", textBox10.Text);
                archive.SetValue("PTO_WeightBrutto", textBox9.Text);

                archive.SetValue("PTO_PR", textBox1.Text);
                archive.SetValue("PTO_PG", textBox2.Text);
                archive.SetValue("PTO_T", textBox11.Text);
                archive.SetValue("PTO_PlaneNumber", textBox4.Text);

                archive.SetValue("People_Razrab", textBox13.Text);
                archive.SetValue("People_Prov", textBox14.Text);
                archive.SetValue("People_Nach", textBox16.Text);
                archive.SetValue("People_Norm", textBox15.Text);
                archive.SetValue("People_Utv", textBox17.Text);

                archive.SetValue("Medium_MediumHot", txtMediumHot.Text);
                archive.SetValue("Medium_MediumCold", txtMediumCold.Text);
                archive.SetValue("Medium_TypeHot", comboBox13.SelectedIndex);
                archive.SetValue("Medium_TypeCold", comboBox14.SelectedIndex);
                archive.SetValue("Medium_MaxThot", textBox7.Text);
                archive.SetValue("Medium_MaxTcold", textBox12.Text);
                archive.SetValue("Medium_MinT", textBox8.Text);

                archive.SetValue("Material_StaticPlate", comboBox5.SelectedIndex);
                archive.SetValue("Material_PriPlate", comboBox6.SelectedIndex);
                archive.SetValue("Material_NaprUp", comboBox7.SelectedIndex);
                archive.SetValue("Material_NaprDown", comboBox8.SelectedIndex);
                archive.SetValue("Material_Rez", comboBox12.SelectedIndex);
                archive.SetValue("Material_Shplilka", comboBox9.SelectedIndex);
                archive.SetValue("Material_Bolt", comboBox11.SelectedIndex);
                archive.SetValue("Material_Gaika", comboBox10.SelectedIndex);
                archive.SetValue("Material_Phlance", comboBox16.SelectedIndex);

                string startupPath = Source_Path + "\\";
                string sId = "1";

                if (File.Exists(startupPath + "id.txt"))
                {
                    using (StreamReader sr = new StreamReader(startupPath + "id.txt"))
                    {
                        sId = sr.ReadLine();
                    }
                }
                else
                {
                    using (StreamWriter swId = new StreamWriter(startupPath + "id.txt"))
                    {
                        swId.Write("1");
                    }
                }

                int id = 0;
                int.TryParse(sId, out id);

                SReportDataBase db = new SReportDataBase();

                if (File.Exists(startupPath + "db.xml"))
                {
                    db.Open(startupPath + "db.xml");
                }

                db.Id = id;
                db.StartupPath = startupPath;
                
                archive.Save(db.Add(dlg.ReportName, dlg.Comment));
                db.Save(startupPath + "db.xml");

                using (StreamWriter sw = new StreamWriter(startupPath + "id.txt"))
                {
                    sw.WriteLine(db.Id.ToString());
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            frmOpen dlg = new frmOpen();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                SReportArchive archive = new SReportArchive();

                // open report
                if (!archive.Open(dlg.Path))
                {
                    MessageBox.Show("Невозможно открыть расчет: " + dlg.ReportName, KitConstant.softwareName, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }

                // change title
                this.Text = string.Format("{0}: {1}", KitConstant.softwareName, dlg.ReportName);

                comboBox1.SelectedIndex = archive.GetIntegerValue("PTO_Type");
                textBox3.Text = archive.GetStringValue("PTO_PlateCount");
                comboBox2.SelectedIndex = archive.GetIntegerValue("PTO_PlateMaterial");
                comboBox3.SelectedIndex = archive.GetIntegerValue("PTO_PlateThickness");
                comboBoxWrapper.SelectedIndex = archive.GetIntegerValue("PTO_WrapperMaterial1");
                comboBox4.SelectedIndex = archive.GetIntegerValue("PTO_WrapperMaterial2");

                textBox5.Text = archive.GetStringValue("PTO_Volume");
                comboBox17.SelectedIndex = archive.GetIntegerValue("PTO_Len");
                textBox10.Text = archive.GetStringValue("PTO_WeightNetto");
                textBox9.Text = archive.GetStringValue("PTO_WeightBrutto");

                textBox1.Text = archive.GetStringValue("PTO_PR");
                textBox2.Text = archive.GetStringValue("PTO_PG");
                textBox11.Text = archive.GetStringValue("PTO_T");
                textBox4.Text = archive.GetStringValue("PTO_PlaneNumber");

                textBox13.Text = archive.GetStringValue("People_Razrab");
                textBox14.Text = archive.GetStringValue("People_Prov");
                textBox16.Text = archive.GetStringValue("People_Nach");
                textBox15.Text = archive.GetStringValue("People_Norm");
                textBox17.Text = archive.GetStringValue("People_Utv");

                txtMediumHot.Text = archive.GetStringValue("Medium_MediumHot");
                txtMediumCold.Text = archive.GetStringValue("Medium_MediumCold");
                comboBox13.SelectedIndex = archive.GetIntegerValue("Medium_TypeHot");
                comboBox14.SelectedIndex = archive.GetIntegerValue("Medium_TypeCold");
                textBox7.Text = archive.GetStringValue("Medium_MaxThot");
                textBox12.Text = archive.GetStringValue("Medium_MaxTcold");
                textBox8.Text = archive.GetStringValue("Medium_MinT");

                comboBox5.SelectedIndex = archive.GetIntegerValue("Material_StaticPlate");
                comboBox6.SelectedIndex = archive.GetIntegerValue("Material_PriPlate");
                comboBox7.SelectedIndex = archive.GetIntegerValue("Material_NaprUp");
                comboBox8.SelectedIndex = archive.GetIntegerValue("Material_NaprDown");
                comboBox12.SelectedIndex = archive.GetIntegerValue("Material_Rez");
                comboBox9.SelectedIndex = archive.GetIntegerValue("Material_Shplilka");
                comboBox11.SelectedIndex = archive.GetIntegerValue("Material_Bolt");
                comboBox10.SelectedIndex = archive.GetIntegerValue("Material_Gaika");
                comboBox16.SelectedIndex = archive.GetIntegerValue("Material_Phlance");
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.S)
            {
                button6_Click(sender, null);
            }
            else if (e.Control && e.KeyCode == Keys.O)
            {
                button7_Click(sender, null);
            }
            else if (e.Control && e.KeyCode == Keys.P)
            {
                button5_Click(sender, null);
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                button3_Click(sender, null);
            }
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox4.Enabled = (comboBox16.SelectedIndex < 3);
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back || e.KeyChar == m_nfi.NumberDecimalSeparator[0]))
            {
                e.Handled = true;
            }
        }

        private void btnMediumHot_Click(object sender, EventArgs e)
        {
            contextMenuStrip1.Show(btnMediumHot, 0, btnMediumHot.Height + 2);
        }

        private void btnMediumCold_Click(object sender, EventArgs e)
        {
            contextMenuStrip2.Show(btnMediumCold, 0, btnMediumCold.Height + 2);
        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            txtMediumHot.Text = e.ClickedItem.Text;
        }

        private void contextMenuStrip2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            txtMediumCold.Text = e.ClickedItem.Text;
        }
    }
}