using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedVerificationProgram
{
    internal class FinalResultMatix
    {
        public int[,] finalResultMatix = new int[18, 3];
        public string[,] finalResultMatixString = new string[18, 3];

        public FinalResultMatix()
        {
            InitiateFinalResultMatix();
        }

        public void InitiateFinalResultMatix()
        {
            for (int i = 0; i < 32; i++)
            {
                finalResultMatix[i, 0] = -1; // Method
                finalResultMatix[i, 1] = -1; // Result
                finalResultMatix[i, 2] = -1; // MDL
            }
            for (int i = 0; i < 32; i++)
            {
                finalResultMatixString[i, 0] = ""; // Method String
                finalResultMatixString[i, 1] = ""; // Result String
                finalResultMatixString[i, 2] = ""; // MDL String
            }
        }
        public void dispose()
        {
            if (this.finalResultMatix != null)
            {
                this.finalResultMatix = null;
            }

            if (this.finalResultMatixString != null)
            {
                this.finalResultMatixString = null;
            }
        }
    }
}
