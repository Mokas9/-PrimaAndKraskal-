using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrimaAndKraskal
{
    public class Edge: IComparable<Edge>
    {
        public int v1 { get; set; }  //1 вершина
        public int v2 { get; set; }  //2 вершина
        public int weight { get; set; }  //вес ребра e = (v1, v2)

        public int CompareTo(Edge obj)
        {
            if (this.weight > obj.weight)
                return 1;
            if (this.weight < obj.weight)
                return -1;
            else
                return 0;
        }

        public Edge(int v1, int v2, int weight)
        {
            this.v1 = v1;
            this.v2 = v2;
            this.weight = weight;
        }
    }

    public class PrimsAlgoritm
    {
        public void Calculate(int numberV, List<Edge> E, List<Edge> MST)
        {
            //неиспользованные рёбра
            List<Edge> notUsedE = new List<Edge>(E);
            //использованные вершины
            List<int> usedV = new List<int>();
            //неиспользованные вершины
            List<int> notUsedV = new List<int>();
           
            for (int i = 0; i < numberV; i++)
            {
                notUsedV.Add(i);
            }

            //выбираем случайную начальную вершину

            Random rand = new Random();

            usedV.Add(rand.Next(0, numberV));

            notUsedV.RemoveAt(usedV[0]);

            while(notUsedV.Count > 0)
            {
                int minE = -1;  //номер наименьшего ребра
                //Поиск наименьшего ребра

                for(int i = 0; i < notUsedE.Count; i++)
                {
                    if ((usedV.IndexOf(notUsedE[i].v1) != -1) && (notUsedV.IndexOf(notUsedE[i].v2) != -1) || (usedV.IndexOf(notUsedE[i].v2) != -1) && (notUsedV.IndexOf(notUsedE[i].v1) != -1))
                    {
                        if (minE != -1)
                        {
                            if (notUsedE[i].weight < notUsedE[minE].weight)
                                minE = i;
                        }
                        else
                            minE = i;
                    }
                }

                //заносим новую вершину в список использованных и удаляем её из списка неиспользованных
                if (usedV.IndexOf(notUsedE[minE].v1) != -1)
                {
                    usedV.Add(notUsedE[minE].v2);
                    notUsedV.Remove(notUsedE[minE].v2);
                }
                else
                {
                    usedV.Add(notUsedE[minE].v1);
                    notUsedV.Remove(notUsedE[minE].v1);
                }
                //заносим новое ребро в дерево и удаляем его из списка неиспользованных
                MST.Add(notUsedE[minE]);
                notUsedE.RemoveAt(minE);
            }
        }
    }

    public class KruskalsAlgoritm
    {
        public void Calculate(int numberV, List<Edge> E, List<Edge> MST)
        {
            List<List<int>> mnoshestvo = new List<List<int>>();
           
            for(int i = 0; i < numberV; i++)
            {
                mnoshestvo.Add(new List<int>() { i });
            }

            foreach(Edge e in E)
            {
                int index1 = -1;
                int index2 = -1;

                for (int i = 0; i < mnoshestvo.Count; i++)
                {
                    for (int j = 0; j < mnoshestvo[i].Count; j++)
                    {
                        if (e.v1 == mnoshestvo[i][j])
                            index1 = i;
                        if (e.v2 == mnoshestvo[i][j])
                            index2 = i;
                    }
                }

                if (!(index1 == index2))
                {
                    for (int i = 0; i < mnoshestvo[index2].Count; i++)
                        mnoshestvo[index1].Add(mnoshestvo[index2][i]);
                    mnoshestvo.Remove(mnoshestvo[index2]);
                    MST.Add(e);
               }
            }
        }
    }
}
