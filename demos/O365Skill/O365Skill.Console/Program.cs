using System;
using System.Collections.Generic;

namespace O365Skill.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            using(O365Skill.Graph.GraphClient graphClient=new Graph.GraphClient())
            {
                //List<string> result = graphClient.GetDocument().Result;
                //System.Console.WriteLine(string.Join("\t", result));

                //System.Console.WriteLine(graphClient.GetUserInfo().Result);
                if (graphClient.CreateAppointment().Result) System.Console.WriteLine("OK");
                else System.Console.WriteLine("KO");
                System.Console.ReadLine();
            }
        }
    }
}
