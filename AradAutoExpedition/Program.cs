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

        static Process GetDNFProcess()
        {
            var process = Process.GetProcessesByName("ARAD");
            if (!process.Any())
            {
                process = Process.GetProcessesByName("DFO");
            }

            if (!process.Any())
            {
                Console.WriteLine("DNF is not running");
                return null;
            }

            return process.First();
        }

        static void RunMacro(RECT rect)
        {
            if (rect.Width == 1600 && rect.Height == 900)
            {
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
            else if (rect.Width == 800 && rect.Height == 600)
            {
                var dungeons = new List<Point>();
                dungeons.Add(new Point(120, 340));
                dungeons.Add(new Point(120, 280));
                dungeons.Add(new Point(120, 220));
                dungeons.Add(new Point(120, 160));

                foreach (var dungeon in dungeons)
                {
                    MouseMoveAndLeftClick(dungeon.X, dungeon.Y);
                    MouseMoveAndLeftClick(410, 440);
                    MouseMoveAndLeftClick(380, 508);
                    MouseMoveAndLeftClick(493, 354);
                    MouseMoveAndLeftClick(290, 520);
                    MouseMoveAndLeftClick(dungeon.X, dungeon.Y);
                    MouseMoveAndLeftClick(490, 505);
                    MouseMoveAndLeftClick(360, 400);
                    MouseMoveAndLeftClick(400, 400);
                }
            }
        }

        static void Main(string[] args)
        {
            var process = GetDNFProcess();

            var pid = process.Id;
            var hWnd = process.MainWindowHandle;
            if (hWnd != IntPtr.Zero)
            {
                GetClientRect(hWnd, out var rect);

                if (
                    !(rect.Width == 1600 && rect.Height == 900) &&
                    !(rect.Width == 800 && rect.Height == 600)
                   )
                {
                    Console.WriteLine("Please set resolution 1600x900 or 800x600");
                    return;
                }

                MoveWindow(hWnd, 0, 0, rect.Width, rect.Height, true);
                Microsoft.VisualBasic.Interaction.AppActivate(pid);
                Thread.Sleep(500);

                RunMacro(rect);
            }
        }
    }
}
