using System;
using System.Runtime.InteropServices;

namespace BrokenEvent.VisualStudioOpener.DarkMagic
{
  /// <summary>
  /// Class containing the IOleMessageFilter thread error-handling functions.
  /// </summary>
  /// <remarks>https://msdn.microsoft.com/en-us/library/ms228772.aspx</remarks>
  public class MessageFilter : IOleMessageFilter
  {
    /// <summary>
    /// Start the filter.
    /// </summary>
    public static void Register()
    {
      IOleMessageFilter newFilter = new MessageFilter();
      IOleMessageFilter oldFilter = null;
      CoRegisterMessageFilter(newFilter, out oldFilter);
    }

    /// <summary>
    /// Done with the filter, close it.
    /// </summary>
    public static void Revoke()
    {
      IOleMessageFilter oldFilter = null;
      CoRegisterMessageFilter(null, out oldFilter);
    }

    /// <summary>
    /// Handle incoming thread requests.
    /// </summary>
    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
    {
      //Return the flag SERVERCALL_ISHANDLED.
      return 0;
    }

    /// <summary>
    /// Thread call was rejected, so try again.
    /// </summary>
    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
      if (dwRejectType == 2)
      // flag = SERVERCALL_RETRYLATER.
      {
        // Retry the thread call immediately if return >=0 & <100.
        return 99;
      }
      // Too busy; cancel call.
      return -1;
    }

    int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
      int dwTickCount, int dwPendingType)
    {
      //Return the flag PENDINGMSG_WAITDEFPROCESS.
      return 2;
    }

    // Implement the IOleMessageFilter interface.
    [DllImport("Ole32.dll", EntryPoint = "CoRegisterMessageFilter")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
  }
}
