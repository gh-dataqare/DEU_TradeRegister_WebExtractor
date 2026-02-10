using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DEU_TradeRegister_WebExtractor
{
  static class ImportsFromWindowsAPI
  {
    // Define constants for FormatMessageW flags
    private const uint FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;
    private const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;
    private const uint FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;


    // Retains the current Z order (ignores the hWndInsertAfter parameter).
    private const uint m_cuiSWP_NOZORDER = 0x0004;


    [DllImport("kernel32.dll", EntryPoint = "GetLastError")]
    private static extern int GetLastError();


    // P/Invoke declaration for FormatMessageW
    [DllImport("kernel32.dll", EntryPoint = "FormatMessageW", CharSet = CharSet.Unicode)]
      private static extern int FormatMessageW(
        uint dwFlags,
        IntPtr lpSource,
        int dwMessageId,
        uint dwLanguageId,
        [Out] StringBuilder lpBuffer, // Use StringBuilder for allocated buffer
        uint nSize,
        IntPtr Arguments // Use IntPtr for arguments, or a specific type if needed
     );
      

    [DllImport("user32.dll", EntryPoint = "SendInput")]
    private static extern uint SendInput(uint nInputs, byte[] inputs, int cbSize);



    [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
    private static extern int SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, uint wFlags);

    [DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", EntryPoint = "SetActiveWindow")]
    private static extern IntPtr SetActiveWindow(IntPtr hWnd);

    [DllImport("user32.dll", EntryPoint = "GetFocus")]
    private static extern IntPtr GetFocus();

    [DllImport("user32.dll", EntryPoint = "AttachThreadInput")]
    private static extern bool AttachThreadInput(int idAttach, int idAttachTo, bool fAttach);

    [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")]
    private static extern int GetCurrentThreadId();

    [DllImport("user32.dll", EntryPoint = "FindWindowW", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

    [DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll", EntryPoint = "SendMessageW", CharSet = CharSet.Unicode)]
    private static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam, uint lParam);



    public static String GetMessageForWindowsErrorCode(int v_iErrorCode)
    {
      String sResult = "";

      StringBuilder lpBuffer = new StringBuilder(12000); // Initial buffer size

      int iCountOfWrittenChars = FormatMessageW(
        FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        IntPtr.Zero, // No specific source for system messages
        v_iErrorCode,
        0, // Default language ID
        lpBuffer,
        (uint)lpBuffer.Capacity,
        IntPtr.Zero // No arguments for system messages
      );

      if (iCountOfWrittenChars > 0)
      {
        sResult = lpBuffer.ToString();
      }
      else
      {
        sResult =
          "Cannot retrieve Windows message text for error id: " +
          v_iErrorCode.ToString() + ".";
      }

      if (sResult.EndsWith("\r\n"))
        sResult = sResult.Substring(0, sResult.Length - 2);

      return sResult;
    }

    public static void SetWindowPosition(
      IntPtr v_iMainWindowHandle,
      int v_iLeft,
      int v_iTop,
      int v_iWidth,
      int v_iHeighе,
      String v_sPrefixOfErrorMessage)
    {
      int iWinAPIResult = SetWindowPos(
        v_iMainWindowHandle, 0,
        v_iLeft, v_iTop, v_iWidth, v_iHeighе,
        m_cuiSWP_NOZORDER);

      if (iWinAPIResult == 0)
      {
        int iLastError = 0;
        String sErrorMessage = "";


        iLastError = GetLastError();

        if (v_sPrefixOfErrorMessage != null &&
            v_sPrefixOfErrorMessage.Length != 0)
        {
          sErrorMessage = sErrorMessage + v_sPrefixOfErrorMessage + "\n";
        }

        sErrorMessage = sErrorMessage +
          "Error in SetWindowPos() function from user32.dll. " +
          "GetLastError() = " + iLastError.ToString() + ". " +
          "Windows error message = {" + GetMessageForWindowsErrorCode(iLastError) + "}.";

        throw new Exception(sErrorMessage);
      }
    }

    public static string GetActiveWindowTitle()
    {
      const int nChars = 256;
      StringBuilder Buff = new StringBuilder(nChars);
      IntPtr handle = GetForegroundWindow();

      if (GetWindowText(handle, Buff, nChars) > 0)
      {
        return Buff.ToString();
      }

      return null;
    }
  }
}
