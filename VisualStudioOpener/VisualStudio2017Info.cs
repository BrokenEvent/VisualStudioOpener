using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;

using EnvDTE;

using Process = System.Diagnostics.Process;
using Thread = System.Threading.Thread;

namespace BrokenEvent.VisualStudioOpener
{
  /// <summary>
  /// Using sample code from
  /// http://www.helixoft.com/blog/creating-envdte-dte-for-vs-2017-from-outside-of-the-devenv-exe.html
  /// </summary>
  internal class VisualStudio2017Info : VisualStudioInfo
  {
    public VisualStudio2017Info(Version version, string description, string path): base(version, description, path) {}

    protected override DTE GetDTE()
    {
      return GetDTE(Path);
    }

    #region Dark Magic

    /// <summary>
    /// Tries to find running VS instance of the specified version. If not, new instance will be launched.
    /// </summary>
    /// <param name="devenvPath">The full path to the devenv.exe.</param>
    /// <returns>DTE instance or null if not found.</returns>
    private static DTE GetDTE(string devenvPath)
    {
      Process[] processes = Process.GetProcessesByName("devenv");
      Process devenvProcess = null;
      foreach (Process process in processes)
        try
        {
          if (process.MainModule.FileName.Equals(devenvPath, StringComparison.InvariantCultureIgnoreCase))
          {
            devenvProcess = process;
            break;
          }
        }
        catch {}

      if (devenvProcess != null)
        return GetDTEFromInstance(devenvProcess.Id, 120);

      return CreateDTEInstance(devenvPath);
    }

    /// <summary>
    /// Creates and returns a DTE instance of specified VS version.
    /// </summary>
    /// <param name="devenvPath">The full path to the devenv.exe.</param>
    /// <returns>DTE instance or null if not found.</returns>
    private static DTE CreateDTEInstance(string devenvPath)
    {
      DTE dte;
      Process proc;

      // start devenv
      ProcessStartInfo procStartInfo = new ProcessStartInfo();
      procStartInfo.Arguments = "-Embedding";
      procStartInfo.CreateNoWindow = true;
      procStartInfo.FileName = devenvPath;
      procStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
      procStartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(devenvPath);

      try
      {
        proc = Process.Start(procStartInfo);
      }
      catch (Exception)
      {
        return null;
      }

      if (proc == null)
        return null;

      // get DTE
      dte = GetDTEFromInstance(proc.Id, 120);

      return dte;
    }

    /// <summary>
    /// Gets the DTE object from any devenv process.
    /// </summary>
    /// <remarks>
    /// After starting devenv.exe, the DTE object is not ready. We need to try repeatedly and fail after the
    /// timeout.
    /// </remarks>
    /// <param name="processId">The process id</param>
    /// <param name="timeout">Timeout in seconds.</param>
    /// <returns>
    /// Retrieved DTE object or null if not found.
    /// </returns>
    private static DTE GetDTEFromInstance(int processId, int timeout)
    {
      DTE res = null;
      DateTime startTime = DateTime.Now;

      while (res == null && DateTime.Now.Subtract(startTime).Seconds < timeout)
      {
        Thread.Sleep(1000);
        res = GetDTE(processId);
      }

      return res;
    }

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

    /// <summary>
    /// Gets the DTE object from any devenv process.
    /// </summary>
    /// <param name="processId">The Process id</param>
    /// <returns>
    /// Retrieved DTE object or null if not found.
    /// </returns>
    private static DTE GetDTE(int processId)
    {
      object runningObject = null;

      IBindCtx bindCtx = null;
      IRunningObjectTable rot = null;
      IEnumMoniker enumMonikers = null;

      try
      {
        Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
        bindCtx.GetRunningObjectTable(out rot);
        rot.EnumRunning(out enumMonikers);

        IMoniker[] moniker = new IMoniker[1];
        IntPtr numberFetched = IntPtr.Zero;
        while (enumMonikers.Next(1, moniker, numberFetched) == 0)
        {
          IMoniker runningObjectMoniker = moniker[0];

          string name = null;

          try
          {
            if (runningObjectMoniker != null)
              runningObjectMoniker.GetDisplayName(bindCtx, null, out name);
          }
          catch (UnauthorizedAccessException)
          {
            // Do nothing, there is something in the ROT that we do not have access to.
          }

          Regex monikerRegex = new Regex(@"!VisualStudio.DTE\.\d+\.\d+\:" + processId, RegexOptions.IgnoreCase);
          if (!string.IsNullOrEmpty(name) && monikerRegex.IsMatch(name))
          {
            Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
            break;
          }
        }
      }
      finally
      {
        if (enumMonikers != null)
          Marshal.ReleaseComObject(enumMonikers);

        if (rot != null)
          Marshal.ReleaseComObject(rot);

        if (bindCtx != null)
          Marshal.ReleaseComObject(bindCtx);
      }

      return runningObject as DTE;
    }

  #endregion
  }
}
