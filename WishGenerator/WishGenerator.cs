using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class WishGenerator
    {
        List<List<string>> wishes;
        List<List<string>> generated;
        List<List<string>> allCombinations;
        public WishGenerator()
        {
            generated = new List<List<string>>();
            allCombinations = new List<List<string>>();
        }
        public List<List<string>> Generated
        {
            get { return generated; }
        }
        public List<List<string>> Wishes
        {
            set { wishes = value; }
        }
        bool isUsed(List<string> wish)
        {
            bool flag = false;
            if (generated.Count != 0)
                foreach (List<string> usedWish in generated)
                    if (usedWish[0] == wish[0] && usedWish[1] == wish[1] && usedWish[2] == wish[2])
                    {
                        flag = true;
                        break;
                    }
            return flag;
        }
        void getAll()
        {
            int n = wishes.Count;
            for (int i = 0; i < n - 2; i++)
                for (int j = i + 1; j < n - 1; j++)
                    for (int k = j + 1; k < n; k++)
                        for (int l = 0; l < wishes[i].Count; l++)
                            for (int m = 0; m < wishes[j].Count; m++)
                                for (int p = 0; p < wishes[k].Count; p++)
                                {
                                    List<string> wish = new List<string>();
                                    wish.Add(wishes[i][l]);
                                    wish.Add(wishes[j][m]);
                                    wish.Add(wishes[k][p]);
                                    allCombinations.Add(wish);
                                }
        }
        public void Generate(int n)
        {
            getAll();
            while (generated.Count != n)
            {
                Random rand = new Random();
                int wishNum = rand.Next() % allCombinations.Count;
                if (!isUsed(allCombinations[wishNum]))
                    generated.Add(allCombinations[wishNum]);
            }
        }
    }
}
