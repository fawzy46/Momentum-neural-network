using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace ML
{
    public partial class Form1 : Form
    {
        network nk = new network();
        string[] nsize, lsize;
        int index = 0;
        double[,] inputs;
        double[] sm;
        double ds;
        double alpha;
        int iteration;
        int it = 0;
        int nr;
        double LR;
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        _Excel.Range range;
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        void initialize_weights(int len1, int len2, int n1)
        {
            Random r = new Random();
            layer ly = new layer();
            for(int i = 0; i < len2 ; i++)                  // i = dest
            {
                for (int k = 0; k < len1 ; k++)            // k = scr
                {
                    weight w = new weight();
                    w.val = 0;//r.NextDouble();
                    w.lscr = n1;
                    w.ldest = n1+1;
                    w.nscr = k;
                    w.ndest = i;
                    nk.ly[n1+1].nl[i].w.Add(w);
                }
            }
        }
        void initialize_nodes(int len1)
        {
            layer ly = new layer();
            for(int i = 0; i < len1 ;i++)
            {
                node nd = new node();
                ly.nl.Add(nd);
            }
            nk.ly.Add(ly);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            nsize = textBox5.Text.Split(' ');
            lsize = textBox6.Text.Split(' ');
            LR = double.Parse(textBox3.Text);
            alpha = double.Parse(textBox1.Text);
            ds = double.Parse(textBox2.Text);
            iteration = int.Parse(textBox4.Text);
            int j = 0;
            for(int i = 0; i < nsize.Length; i++)
            {
                int x = int.Parse(nsize[i]);
                for (int k = 0; k < x ; k++)
                {
                    initialize_nodes(int.Parse(lsize[j]));
                    j++;
                }
            }
            for(int i = 0; i < nk.ly.Count - 1; i++)
            {
                initialize_weights(nk.ly[i].nl.Count, nk.ly[i+1].nl.Count, i);
            }
            while (it < iteration)
            {
                for (int s = 0; s < nr; s++)
                {
                    forward_ph();
                }
                index = 0;
                it++;
                disp_weights();
            }
        }
        void disp_weights()
        {
            for(int i = 1; i < nk.ly.Count;i++)
            {
                for(int k = 0; k < nk.ly[i].nl.Count;k++)
                {
                    for(int w = 0; w < nk.ly[i].nl[k].w.Count;w++)
                    {
                        MessageBox.Show(nk.ly[i].nl[k].w[w].val.ToString());
                    }
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            wb = excel.Workbooks.Open("C:\\Users\\fawzy ahmed\\source\\repos\\ML\\ML\\bin\\Debug\\s.xlsx");
            ws = excel.Worksheets[1];
            range = ws.UsedRange;
            int row = range.Rows.Count;
            int col = range.Columns.Count;
            nr = row;
            inputs = new double[row, col];
            int l = 0, j = 0;
            for (int i = 1; i <= row; i++)
            {
                l = 0;
                for(int k = 1; k <= col; k++)
                {
                    inputs[j, l] = range.Cells[i, k].Value2;
                    l++;
                }
                j++;
            }
            wb.Close();
            excel.Quit();
        }

        void forward_ph()
        {

            sm = new double[int.Parse(lsize[0])];
            for(int i = 0; i < int.Parse(lsize[0]) ;i++)
            {
                sm[i] = inputs[index,i];
            }
            index++;
            MessageBox.Show(index.ToString());
            for(int i = 0; i < nk.ly[0].nl.Count;i++)
            {
                nk.ly[0].nl[i].input = sm[i];
            }
            for(int i = 1; i < nk.ly.Count;i++)
            {
                for(int k = 0; k < nk.ly[i].nl.Count;k++)
                {
                    for(int j = 0; j < nk.ly[i].nl[k].w.Count; j++)
                    {
                        int nscr = nk.ly[i].nl[k].w[j].nscr;
                        nk.ly[i].nl[k].input += (nk.ly[i].nl[k].w[j].val) * (nk.ly[i-1].nl[nscr].input);
                    }
                    nk.ly[i].nl[k].input = (1/(1+Math.Pow(Math.E, -nk.ly[i].nl[k].input)));
                }
            }
            if (nk.ly[nk.ly.Count-1].nl[0].input != ds)
            {
                backward_ph();
            }
        }
        void backward_ph()
        {
            for(int i = 0; i < nk.ly[nk.ly.Count-1].nl.Count; i++)
            {
                nk.ly[nk.ly.Count - 1].nl[i].error = (nk.ly[nk.ly.Count - 1].nl[i].input) * (1 - nk.ly[nk.ly.Count - 1].nl[i].input) * (ds - nk.ly[nk.ly.Count - 1].nl[i].input); 
            }
            for(int i = nk.ly.Count - 2; i > 0;i--)
            {
                for(int k = 0; k < nk.ly[i].nl.Count; k++)
                {
                    double temp = 0;
                    for(int j = 0; j < nk.ly[i+1].nl.Count; j++)
                    {
                        for(int w = 0; w < nk.ly[i+1].nl[j].w.Count;w++)
                        {
                            if(nk.ly[i + 1].nl[j].w[w].lscr == i && nk.ly[i + 1].nl[j].w[w].nscr == k)
                            {
                                temp += nk.ly[i + 1].nl[j].error * nk.ly[i + 1].nl[j].w[w].val;
                                break;
                            }
                        }

                    }
                    nk.ly[i].nl[k].error = (nk.ly[i].nl[k].input) * (1 - nk.ly[i].nl[k].input) * temp;
                }
            }
            update_weight();
        }
        void update_weight()
        {
            for(int i = 1; i < nk.ly.Count; i++)
            {
                for(int k = 0; k < nk.ly[i].nl.Count;k++)
                {
                    for(int w = 0; w < nk.ly[i].nl[k].w.Count;w++)
                    {
                        int scr = nk.ly[i].nl[k].w[w].nscr;
                        nk.ly[i].nl[k].w[w].val = (alpha * nk.ly[i].nl[k].w[w].val) + (LR * nk.ly[i].nl[k].error * nk.ly[i-1].nl[scr].input);
                    }
                }
            }
        }
    }
    public class layer
    {
        public List<node> nl = new List<node>();
    }
    public class node
    {
        public double input, output, error;
        public List<weight> w = new List<weight>();
    }
    public class network
    {
        public List<layer> ly = new List<layer>();
    }
    public class weight
    {
        public double val;
        public int lscr, ldest, nscr, ndest;
    }
}
