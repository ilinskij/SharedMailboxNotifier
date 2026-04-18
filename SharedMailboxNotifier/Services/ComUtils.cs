using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SharedMailboxNotifier.Services
{
    internal static class ComHelper
    {
        internal static void SafeComRelease(object comObject)
        {
            if (comObject == null)
                return;

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("[COM] Failed to properly release the "
                    + ex.GetType().Name + " COM object: " + ex.Message
                    + Environment.NewLine + ex.StackTrace);
            }
        }
    }
}
