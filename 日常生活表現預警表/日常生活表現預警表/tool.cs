using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 日常生活表現預警表
{
    static class tool
    {
        static public FISCA.UDT.AccessHelper _A = new FISCA.UDT.AccessHelper();

        static public FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();


        static public Dictionary<string, int> GetColumnTitle(bool Is換算, bool Is相抵)
        {
            Dictionary<string, int> dic = new Dictionary<string, int>();
            dic.Add("班級", 0);
            dic.Add("學號", 1);
            dic.Add("座號", 2);
            dic.Add("姓名", 3);

            dic.Add("本期大功", 4);
            dic.Add("本期小功", 5);
            dic.Add("本期嘉獎", 6);
            dic.Add("本期大過", 7);
            dic.Add("本期小過", 8);
            dic.Add("本期警告", 9);

            dic.Add("歷年大功", 10);
            dic.Add("歷年小功", 11);
            dic.Add("歷年嘉獎", 12);
            dic.Add("歷年大過", 13);
            dic.Add("歷年小過", 14);
            dic.Add("歷年警告", 15);

            dic.Add("歷年銷過大過", 16);
            dic.Add("歷年銷過小過", 17);
            dic.Add("歷年銷過警告", 18);

            if (Is換算)
            {
                if (Is相抵)
                {
                    dic.Add("歷年功過換算大功", 19);
                    dic.Add("歷年功過換算小功", 20);
                    dic.Add("歷年功過換算嘉獎", 21);
                    dic.Add("歷年功過換算大過", 22);
                    dic.Add("歷年功過換算小過", 23);
                    dic.Add("歷年功過換算警告", 24);

                    dic.Add("歷年功過相抵大功", 25);
                    dic.Add("歷年功過相抵小功", 26);
                    dic.Add("歷年功過相抵嘉獎", 27);
                    dic.Add("歷年功過相抵大過", 28);
                    dic.Add("歷年功過相抵小過", 29);
                    dic.Add("歷年功過相抵警告", 30);
                }
                else
                {
                    dic.Add("歷年功過換算大功", 19);
                    dic.Add("歷年功過換算小功", 20);
                    dic.Add("歷年功過換算嘉獎", 21);
                    dic.Add("歷年功過換算大過", 22);
                    dic.Add("歷年功過換算小過", 23);
                    dic.Add("歷年功過換算警告", 24);
                }
            }
            else
            {
                if (Is相抵)
                {
                    dic.Add("歷年功過相抵大功", 19);
                    dic.Add("歷年功過相抵小功", 20);
                    dic.Add("歷年功過相抵嘉獎", 21);
                    dic.Add("歷年功過相抵大過", 22);
                    dic.Add("歷年功過相抵小過", 23);
                    dic.Add("歷年功過相抵警告", 24);
                }
            }

            return dic;
        }
    }
}
