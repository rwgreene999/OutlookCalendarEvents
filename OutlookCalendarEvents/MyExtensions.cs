using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalendarEvents
{
    public static class MyExtensions
    {

        public static int TryParse(this string str, int def)
        {
            int thisInt;
            if (int.TryParse(str, out thisInt))
            {
                return thisInt;
            }
            return def;
        }
        public static double TryParse(this string str, double def)
        {
            double result;
            if (double.TryParse(str, out result))
            {
                return result;
            }
            return def;
        }



        public static string ToHexString(this string input)
        {
            var bytes = Encoding.ASCII.GetBytes(input);
            var hexString = BitConverter.ToString(bytes);

            hexString = hexString.Replace("-", " ");

            if (hexString.Contains("0D 0A")) hexString = hexString.Replace("0D 0A", "0D 0A\n");
            else if (hexString.Contains("0D")) hexString = hexString.Replace("0D", "0D\n");
            else if (hexString.Contains("0A")) hexString = hexString.Replace("0A", "0A\n");

            hexString = hexString.Replace("\n ", "\n");

            return hexString;
        }




        // semaphore stuff 
        public static void Block(this System.Threading.ManualResetEvent semaphore)
        {
            semaphore.Reset();
        }
        public static void Release(this System.Threading.ManualResetEvent semaphore)
        {
            semaphore.Set();
        }
        public static void Wait(this System.Threading.ManualResetEvent semaphore)
        {
            semaphore.WaitOne();
        }
        public static void WaitWithTimeout(this System.Threading.ManualResetEvent semaphore, TimeSpan timeSpan)
        {
            semaphore.WaitOne(timeSpan);
        }

    }
}
