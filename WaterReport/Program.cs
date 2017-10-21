using System;
using System.Configuration;
namespace WaterReport
{
    class Program
    {
        static void Main(string[] args)
        {
            /***
             *  接收参数
             *  @param  [0] 类型
             *  @param  [1] 时间
             **/
           
            
            //args = new String[] { "2017-10-20", "M" };
            if (args.Length==2||args.Length==0)
            {
               if (args.Length == 0) args = new String[] {DateTime.Now.ToString(), "D" };
               // if (args.Length == 0) args = new String[] { "2017-09-20", "Y" };
                try
                {
                    DateTime time =  DateTime.Parse(args[0]);
                    string when = args[1].ToUpper();
                    DateTime when1=time;
                    DateTime when2=DateTime.Now;
                    switch (when)
                    {
                        case "Y":
                            when1 = new DateTime(time.Year, 1, 1,0,0,0);
                            when2 = new DateTime(time.Year, 12, 31,23,59,59);
                            break;
                        case "M":
                            when1= time.AddDays(1 - time.Day);  //本月月初
                            when1 = new DateTime(when1.Year, when1.Month, 1, 0, 0, 0);
                            when2 = when1.AddMonths(1).AddDays(-1);//本月月末
                            when2 = new DateTime(when2.Year, when2.Month, when2.Day, 23, 59, 59);
                            break;
                        case "D":
                            when1 = new DateTime(time.Year, time.Month,time.Day, 0, 0, 0);
                            when2 = new DateTime(time.Year, time.Month, time.Day, 23, 59, 59);
                            break;
                        case "W":
                            when1 = time.AddDays(1 - Convert.ToInt32(time.DayOfWeek.ToString("d")));  //本周周一
                            when1 = new DateTime(when1.Year, when1.Month, when1.Day, 0, 0, 0);
                            when2 =  when1.AddDays(6);  //本周周日
                            when2 = new DateTime(when2.Year, when2.Month, when2.Day, 23, 59, 59);
                            break;
                        default://默认是给定时间的天统计
                            when1 = new DateTime(time.Year, time.Month, time.Day, 0, 0, 0);
                            when2 = new DateTime(time.Year, time.Month, time.Day, 23, 59, 59);
                            when = "D";
                            break;
                    }
                    Console.WriteLine(string.Format("{0}-{1}      {2}统计", when1, when2, when));
                    Service.Exists(when1, when2, when);
                }
                catch (Exception e)
                {
                    Console.WriteLine(String.Format("错误{0}",e));
               //     throw e;
                }
               // Console.ReadLine();
            }
        }
    }
}
