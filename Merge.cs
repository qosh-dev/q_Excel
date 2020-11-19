using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel
{
    public class Merge
    {
        public string val;
        public int indStart;
        public int indStop;
        public Merge(string val, int indStart, int indStop)
        {
            this.val = val;
            this.indStart = indStart;
            this.indStop = indStop;
        }

        public static Merge[] buildMerge(string[] data, int step)
        {
            var result = new List<Merge>();
            var list = new List<Merge>();
            int Start = 2;
            int Step = step;
            list.Add(new Merge(data[0], 1, Step));
            Start += Step - 1;
            for (int i = 1; i < data.Count(); i++)
            {
                list.Add(
                    new Merge(data[i], Start, Start + Step - 1)
                );
                Start += Step;
            }
            return list.ToArray();
        }
    }

}
