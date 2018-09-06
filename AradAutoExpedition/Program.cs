using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;

namespace AradAutoExpedition
{
    class Program
    {
        static InputSimulator simulator = new InputSimulator();

        [DllImport("user32.dll")]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        static double GetAbsoluteX(int x) =>
            x * 65535 / Screen.PrimaryScreen.Bounds.Width;

        static double GetAbsoluteY(int y) =>
            y * 65535 / Screen.PrimaryScreen.Bounds.Height;

        static void PressKey(VirtualKeyCode keyCode)
        {
            simulator.Keyboard.KeyDown(keyCode);
            Thread.Sleep(100);
            simulator.Keyboard.KeyUp(keyCode);
        }

        static void MouseLeftClick()
        {
            simulator.Mouse.LeftButtonDown();
            Thread.Sleep(100);
            simulator.Mouse.LeftButtonUp();
            Thread.Sleep(100);
        }

        static void MouseMoveAndLeftClick(int x, int y)
        {
            simulator.Mouse.MoveMouseTo(GetAbsoluteX(x), GetAbsoluteY(y));
            Thread.Sleep(100);
            MouseLeftClick();
        }

        static void Main(string[] args)
        {
            var pid = Process.GetProcessesByName("ARAD").First().Id;
            var hWnd = Process.GetProcessesByName("ARAD").First().MainWindowHandle;
            if (hWnd != IntPtr.Zero)
            {
                GetClientRect(hWnd, out var rect);

                if (rect.Width != 1600 || rect.Height != 900)
                {
                    Console.WriteLine("アラド戦記の解像度を1600x900に設定して下さい。");
                    return;
                }

                MoveWindow(hWnd, 0, 0, rect.Width, rect.Height, true);
            }

            Microsoft.VisualBasic.Interaction.AppActivate(pid);

            Thread.Sleep(500);

            var dungeons = new List<Point>();
            dungeons.Add(new Point(400, 507));
            dungeons.Add(new Point(400, 417));
            dungeons.Add(new Point(400, 327));
            dungeons.Add(new Point(400, 218));

            foreach (var dungeon in dungeons)
            {
                MouseMoveAndLeftClick(dungeon.X, dungeon.Y);
                MouseMoveAndLeftClick(814, 666);
                MouseMoveAndLeftClick(783, 760);
                MouseMoveAndLeftClick(940, 530);
                MouseMoveAndLeftClick(631, 786);
                MouseMoveAndLeftClick(dungeon.X, dungeon.Y);
                MouseMoveAndLeftClick(935, 763);
                MouseMoveAndLeftClick(748, 603);
                MouseMoveAndLeftClick(800, 600);
            }
        }
    }
}
