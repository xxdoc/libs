using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            MemMapFile mmf = new MemMapFile();
            if(!mmf.CreateMemMapFile("DAVES_VFILE", 20)){
                Console.Write("Failed to create vfile");
                return;
            }

            byte[] b = new byte[20];
            if (!mmf.ReadFile(b, 5))
            {
                Console.Write("Failed to read");
            }

            for (int i = 0; i < 5; i++)
            {
                Console.Write("b[" + i + "]=" + ((int)b[i]).ToString("X") + " ");
                b[i] = (byte)(0x41 + i);
            }
            Console.Write("\n");

            if (!mmf.WriteFile(b))
            {
                Console.WriteLine("Failed to write file");
            }

            
            if (!mmf.ReadFile(b, 5))
            {
                Console.Write("Failed to read");
            }

            for (int i = 0; i < 5; i++)
            {
                Console.Write("b[" + i + "]=" + ((int)b[i]).ToString("X") + " ");
            }

            Console.WriteLine("One more read now..\n");
            Console.ReadKey();

            if (!mmf.ReadFile(b, 5))
            {
                Console.Write("Failed to read");
            }

            for (int i = 0; i < 5; i++)
            {
                Console.Write("b[" + i + "]=" + ((int)b[i]).ToString("X") + " ");
            }

            Console.ReadKey(); 

        }
    }
}
