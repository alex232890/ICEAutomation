﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace ImageComposeEditorAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args[0] == "compose") {
                var composeApp = new ComposeAppService();
                Console.WriteLine("composing...");
                composeApp.Compose(args.Skip(1).ToArray(), m => Console.WriteLine(m), i => drawTextProgressBar(1, 100));
            }
            else if(args[0] == "advanced-compose")
            {
                var composeApp = new AdvancedComposeAppService();
                Console.WriteLine("composing...");
                composeApp.Compose(args.Skip(1).ToArray(), m => Console.WriteLine(m), i => drawTextProgressBar(1, 100));
            }
            else if(args[0] == "process") 
            {
                var composeApp = new ComposeAppService();
                Console.WriteLine("process...");
                var num = args.Length > 1 ? int.Parse(args[1]) : 3;
                var extension = args.Length > 2 ? args[2] : "*.JPG";
                if (args.Length > 3)
                    Directory.SetCurrentDirectory(args[3]);
                var files = GroupFiles(extension, num);
                int total = files.Count;
                int count = 0;
                foreach (var item in files)
                {
                    count++;
                    Console.WriteLine(string.Format("composing {0} of {1}....", count, total));
                    composeApp.Compose(item, m => Console.WriteLine(m), i => drawTextProgressBar(i, 100));
                }
            }
            Console.WriteLine("Finished.");
        }

        private static List<string[]> GroupFiles(string extension, int groupNum)
        {
            string[] filePaths = Directory.GetFiles(Directory.GetCurrentDirectory(), extension, SearchOption.TopDirectoryOnly);

            return filePaths.Select((value, index) => new { value, index })
                    .GroupBy(x => x.index / groupNum, x => Path.GetFileName(x.value)).Select(g => g.ToArray()).ToList();
        }

        private static void drawTextProgressBar(int progress, int total)
        {
            ////draw empty progress bar
            //Console.CursorLeft = 0;
            //Console.Write("["); //start
            //Console.CursorLeft = 32;
            //Console.Write("]"); //end
            //Console.CursorLeft = 1;
            //float onechunk = 30.0f / total;

            ////draw filled part
            //int position = 1;
            //for (int i = 0; i < onechunk * progress; i++)
            //{
            //    Console.BackgroundColor = ConsoleColor.Gray;
            //    Console.CursorLeft = position++;
            //    Console.Write(" ");
            //}

            ////draw unfilled part
            //for (int i = position; i <= 31; i++)
            //{
            //    Console.BackgroundColor = ConsoleColor.Green;
            //    Console.CursorLeft = position++;
            //    Console.Write(" ");
            //}

            ////draw totals
            //Console.CursorLeft = 35;
            //Console.BackgroundColor = ConsoleColor.Black;
            Console.CursorLeft = 15;
            Console.Write(progress.ToString() + " of " + total.ToString() + "    "); //blanks at the end remove any excess
        }
    }
}
