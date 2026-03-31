using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AutomatedVerificationProgram
{
    internal class ReportTestingOrganizationDetction
    {
        public int detectFun(string report_no)
        {
            bool isTwSgs = false;
            bool isCnSgs = false;
            bool isJpSgs = false;
            bool isCti = false;
            bool isKotiti = false;
            bool isKrSgs = false;
            bool isMySgs = false;
            bool isKrIntertek = false;
            bool isTwIntertek = false;

            isTwSgs = Regex.IsMatch(report_no, @"^[E]{1}[TK]{1}[RA]{1}");
            isCnSgs = Regex.IsMatch(report_no, @"^[NHSXCTW]{1}[GKHZMASU]{1}[BGAXNTOH]{1}[EMP]{1}[CL]{1}");
            isJpSgs = Regex.IsMatch(report_no, @"^[J]{1}[EP]{1}");
            isCti = Regex.IsMatch(report_no, @"^[A]{1}[0-9]{1}");
            isKotiti = false; //目前沒有找到規則
            isKrSgs = Regex.IsMatch(report_no, @"^F690101");
            isMySgs = Regex.IsMatch(report_no, @"^[C]{1}[PR]{1}[SP]{1}[ASG]{1}");
            isKrIntertek = Regex.IsMatch(report_no, @"^[R]{1}[T]{1}");
            isTwIntertek = Regex.IsMatch(report_no, @"^[T]{1}[W]{1}[N]{1}[C]{1}");
            

            if (isTwSgs)
            {
                return 0;
            }
            else if (isCnSgs)
            {
                return 1;
            }
            else if (isJpSgs)
            {
                return 2;
            }
            else if (isCti)
            {
                return 3;
            }
            else if (isKotiti)
            {
                return 4;
            }
            else if (isKrSgs)
            {
                return 5;
            }
            else if (isMySgs)
            {
                return 6;
            }
            else if (isKrIntertek)
            {
                return 7;
            }
            else if (isTwIntertek)
            {
                return 8;
            }
            else
            {
                return 9;
            }    
        }
    }
}
