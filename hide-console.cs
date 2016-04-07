using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace HelloWorldApplication
{
   class HelloWorld
   {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);


    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    static void Main(string[] args)
    {
        Console.Title = "ConsoleApplication1";

        IntPtr h=FindWindow(null, "ConsoleApplication
1");

        ShowWindow(h, 0); // 0 = hide

        Form f = new Form();

        f.ShowDialog();

        ShowWindow(h, 1); // 1 = show

      }
   }
}
