using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Collections.Specialized.BitVector32;

namespace 发票助手
{
    internal class UtilTool
    {
        public static String ConvertToChinese(Decimal number)
        {
            var s = number.ToString("#L#E#D#C#K#E#D#C#J#E#D#C#I#E#D#C#H#E#D#C#G#E#D#C#F#E#D#C#.0B0A");
            var d = Regex.Replace(s, @"((?<=-|^)[^1-9]*)|((?'z'0)[0A-E]*((?=[1-9])|(?'-z'(?=[F-L\.]|$))))|((?'b'[F-L])(?'z'0)[0A-L]*((?=[1-9])|(?'-z'(?=[\.]|$))))", "${b}${z}");
            var r = Regex.Replace(d, ".", m => "负元空零壹贰叁肆伍陆柒捌玖空空空空空空空分角拾佰仟万亿兆京垓秭穰"[m.Value[0] - '-'].ToString());
            return r;
        }

        private static readonly Dictionary<char, int> ChineseNumberMap = new Dictionary<char, int>
    {
        { '零', 0 }, { '壹', 1 }, { '贰', 2 }, { '叁', 3 },
        { '肆', 4 }, { '伍', 5 }, { '陆', 6 }, { '柒', 7 },
        { '捌', 8 }, { '玖', 9 }
    };

        private static readonly Dictionary<char, decimal> UnitMap = new Dictionary<char, decimal>
    {
        { '元', 1m }, { '角', 0.1m }, { '分', 0.01m },
        { '拾', 10m }, { '佰', 100m }, { '仟', 1000m },
        { '万', 10000m }, { '亿', 100000000m }
    };

        public static decimal ConvertChineseToDecimal(string chineseAmount)
        {
            decimal result = 0m;
            decimal section = 0m;
            decimal tempNum = 0m;
            decimal lastUnit = 1m;

            foreach (char c in chineseAmount)
            {
                if (ChineseNumberMap.ContainsKey(c))
                {
                    tempNum = ChineseNumberMap[c];
                }
                else if (UnitMap.ContainsKey(c))
                {
                    decimal unit = UnitMap[c];

                    if (unit == 1m || unit == 0.1m || unit == 0.01m)
                    {
                        section += tempNum * unit;
                    }
                    else
                    {
                        section += tempNum * unit;
                        if (unit >= 10000)
                        {
                            result += section * unit;
                            section = 0m;
                        }
                    }

                    tempNum = 0m;
                    lastUnit = unit;
                }
            }

            result += section;


            return result;
        }

    }
}
