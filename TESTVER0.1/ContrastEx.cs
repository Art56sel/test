using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TESTVER0._1
{
    internal class ContrastEx
    {
        private static double LevenshteinDistance(string s, string t)
        {
            if (string.IsNullOrEmpty(s))
            {
                return string.IsNullOrEmpty(t) ? 0 : t.Length;
            }

            if (string.IsNullOrEmpty(t))
            {
                return s.Length;
            }

            int[,] d = new int[s.Length + 1, t.Length + 1];

            for (int i = 0; i <= s.Length; i++)
            {
                d[i, 0] = i;
            }

            for (int j = 0; j <= t.Length; j++)
            {
                d[0, j] = j;
            }

            for (int i = 1; i <= s.Length; i++)
            {
                for (int j = 1; j <= t.Length; j++)
                {
                    int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }

            return d[s.Length, t.Length];
        }
        public void FilterAndDisplaySimilarProducts(string searchWord, DataGridView dataGridView1, DataGridView dataGridView2, DataGridView dataGridView3)
        {
            FilterDataGridViewBySearchWord(searchWord, dataGridView1);
            FilterDataGridViewBySearchWord(searchWord, dataGridView2);
            FilterDataGridViewBySearchWord(searchWord, dataGridView3);
        }
        private void FilterDataGridViewBySearchWord(string searchWord, DataGridView dataGridView)
        {
            // Проходим по всем строкам DataGridView
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                bool containsSearchWord = false;

                // Проходим по всем ячейкам строки
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                    {
                        string cellValue = cell.Value.ToString();

                        // Проверяем, содержит ли значение ячейки заданное слово
                        if (cellValue.Contains(searchWord))
                        {
                            containsSearchWord = true;
                            break;
                        }
                    }
                }

                // Скрываем строку, если в ней не найдено заданное слово
                row.Visible = containsSearchWord;
            }
        }

          
        public void SearchSimilarItemInTable(DataGridView dataGridView, string searchItem)
        {
           
                List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        string itemInTable = row.Cells[0].Value.ToString();
                        int distance = (int)LevenshteinDistance(searchItem, itemInTable);
                        if (distance > 2) // Порог сходства
                        {
                            rowsToRemove.Add(row);
                        }
                    }
                }

                foreach (DataGridViewRow rowToRemove in rowsToRemove)
                {
                    dataGridView.Rows.Remove(rowToRemove);
                }
            }
        }
    }
