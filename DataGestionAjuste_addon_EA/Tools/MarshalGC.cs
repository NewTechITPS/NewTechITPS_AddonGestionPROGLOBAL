using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace REDFARM.Addons.Tools
{
    public class MarshalGC
    {
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, UIntPtr minimumWorkingSetSize, UIntPtr maximumWorkingSetSize);

        public static void ReleaseComObject(object? obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
                GC.Collect();
            }
        }

        public static void ReleaseComObjects(params object?[] objects)
        {
            if (ReleaseValidComObjects(objects))
                GC.Collect();
        }

        public static void ReleaseComObjectsAndMinimize(params object?[] objects)
        {
            if (ReleaseValidComObjects(objects))
                MinimizeMemory();
        }

        public static void MinimizeMemory()
        {
            GC.Collect(GC.MaxGeneration);
            GC.WaitForPendingFinalizers();
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, (UIntPtr)0xFFFFFFFF, (UIntPtr)0xFFFFFFFF);
        }

        private static bool ReleaseValidComObjects(params object?[] objects)
        {
            bool release = false;
            foreach (var obj in objects)
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);

                    release = true;
                }
            }

            return release;
        }
    }
}
