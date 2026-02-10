using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;


namespace DEU_TradeRegister_WebExtractor
{
  public static class KeyboardEmulator
  {
    public static UInt16 m_uiVirtualKey_for_VK_BACK = 0x00000008;   // Backspace key
    public static UInt16 m_uiVirtualKey_for_VK_TAB = 0x00000009;    // Tab key
    public static UInt16 m_uiVirtualKey_for_VK_CLEAR = 0x0000000C;  // Clear key
    public static UInt16 m_uiVirtualKey_for_VK_RETURN = 0x0000000D; // Enter key
    // public static UInt16 m_uiVirtualKey_for_VK_SHIFT = 0x00000010;  // Shift key

    public static UInt16 m_uiVirtualKey_for_VK_PAUSE = 0x00000013;    // Pause key
    public static UInt16 m_uiVirtualKey_for_VK_CAPITAL = 0x00000014;  // Caps lock key

    public static UInt16 m_uiVirtualKey_for_VK_ESCAPE = 0x0000001B;   // Esc key

    public static UInt16 m_uiVirtualKey_for_VK_SPACE = 0x00000020;    // Space key
    public static UInt16 m_uiVirtualKey_for_VK_PRIOR = 0x00000021;    // Page up key
    public static UInt16 m_uiVirtualKey_for_VK_NEXT = 0x00000022;     // Page down key
    public static UInt16 m_uiVirtualKey_for_VK_END = 0x00000023;      // End key
    public static UInt16 m_uiVirtualKey_for_VK_HOME = 0x00000024;     // Home key

    public static UInt16 m_uiVirtualKey_for_VK_LEFT = 0x00000025;   // Left arrow key
    public static UInt16 m_uiVirtualKey_for_VK_UP = 0x00000026;     // Up arrow key
    public static UInt16 m_uiVirtualKey_for_VK_RIGHT = 0x00000027;  // Right arrow key
    public static UInt16 m_uiVirtualKey_for_VK_DOWN = 0x00000028;   // Down arrow key

    public static UInt16 m_uiVirtualKey_for_VK_SELECT = 0x00000029;   // Select key
    public static UInt16 m_uiVirtualKey_for_VK_PRINT = 0x0000002A;    // Print key
    public static UInt16 m_uiVirtualKey_for_VK_EXECUTE = 0x0000002B;  // Execute key
    public static UInt16 m_uiVirtualKey_for_VK_SNAPSHOT = 0x0000002C; // Print screen key
    public static UInt16 m_uiVirtualKey_for_VK_INSERT = 0x0000002D;   // Insert key
    public static UInt16 m_uiVirtualKey_for_VK_DELETE = 0x0000002E;   // Delete key
    public static UInt16 m_uiVirtualKey_for_VK_HELP = 0x0000002F;     // Help key


    public static UInt16 m_uiVirtualKey_for_VK_0 = 0x00000030;  // '0' key
    public static UInt16 m_uiVirtualKey_for_VK_1 = 0x00000031;  // '1' key
    public static UInt16 m_uiVirtualKey_for_VK_2 = 0x00000032;  // '2' key
    public static UInt16 m_uiVirtualKey_for_VK_3 = 0x00000033;  // '3' key
    public static UInt16 m_uiVirtualKey_for_VK_4 = 0x00000034;  // '4' key
    public static UInt16 m_uiVirtualKey_for_VK_5 = 0x00000035;  // '5' key
    public static UInt16 m_uiVirtualKey_for_VK_6 = 0x00000036;  // '6' key
    public static UInt16 m_uiVirtualKey_for_VK_7 = 0x00000037;  // '7' key
    public static UInt16 m_uiVirtualKey_for_VK_8 = 0x00000038;  // '8' key
    public static UInt16 m_uiVirtualKey_for_VK_9 = 0x00000039;  // '9' key

    [DllImport("kernel32.dll", EntryPoint = "GetLastError")]
    private static extern int GetLastError();

    [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")]
    public static extern int GetCurrentThreadId();

    [DllImport("user32.dll", EntryPoint = "SendInput")]
    private static extern uint SendInput(uint nInputs, byte[] inputs, int cbSize);

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

    [DllImport("user32.dll", EntryPoint = "FindWindowW", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

    [DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    // See documentation at Microsoft's web site:
    //   "https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes".
    //
    public static void EmulatePressOfKeyOnKeyboard_ByVirtualKeyCode(UInt16 v_uiVirtualKeyCode)
    {
      UInt16 uiVirtualKeyCode = v_uiVirtualKeyCode;
      byte uchLowByte = (byte)(uiVirtualKeyCode & 0x00FF);
      byte uchHighByte = (byte)((uiVirtualKeyCode >> 8) & 0x00FF);

      byte[] achArrayFor2StructuresOfTypeINPUT = null;
      int iSizeOfINPUTStructure = 0;

      if (Environment.Is64BitProcess)
      {

        // An array for pressing and unpressing of '5' button on a keyboard.
        byte[] achArray_ForChar_5_64bitWin =
        { 
            // KEY_DOWN for the button with code v_uiVirtualKeyCode.
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x35, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,

            // KEY_UP for the button with code v_uiVirtualKeyCode.
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x35, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
          };

        achArray_ForChar_5_64bitWin[8] = uchLowByte;
        achArray_ForChar_5_64bitWin[9] = uchHighByte;
        achArray_ForChar_5_64bitWin[48] = uchLowByte;
        achArray_ForChar_5_64bitWin[49] = uchHighByte;

        achArrayFor2StructuresOfTypeINPUT = achArray_ForChar_5_64bitWin;

        iSizeOfINPUTStructure = 40;
      }
      else
      {
        // An array for pressing and unpressing of '5' button on a keyboard.
        byte[] achArray_ForChar_5_32bitWin =
        {
            // KEY_DOWN for the button with code v_uiVirtualKeyCode.
            0x01, 0x00, 0x00, 0x00, 0x35, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,

            // KEY_UP for the button with code v_uiVirtualKeyCode.
            0x01, 0x00, 0x00, 0x00, 0x35, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
          };

        achArray_ForChar_5_32bitWin[4] = uchLowByte;
        achArray_ForChar_5_32bitWin[5] = uchHighByte;
        achArray_ForChar_5_32bitWin[32] = uchLowByte;
        achArray_ForChar_5_32bitWin[33] = uchHighByte;

        achArrayFor2StructuresOfTypeINPUT = achArray_ForChar_5_32bitWin;

        iSizeOfINPUTStructure = 28;
      }

      UInt32 uiResultCode = SendInput(2, achArrayFor2StructuresOfTypeINPUT, iSizeOfINPUTStructure);

      if (uiResultCode != 2)
      {
        int iLastError = GetLastError();

        throw new Exception("Error in SendInput() function fom user32.dll. " +
          "GetLastError() = " + iLastError.ToString() + ", " +
          "v_uiVirtualKeyCode = 0x" + v_uiVirtualKeyCode.ToString("X4") + ".");
      }
    }

    private static void EmulatePressOfKeyOnKeyboard_ByUnicodeString(String v_sUnicodeString)
    {
      for (int iCharPos = 0; iCharPos < v_sUnicodeString.Length; iCharPos++)
      {
        Char chUnicodeChar = v_sUnicodeString[iCharPos];

        UInt16 uiUnicodeChar = (UInt16)chUnicodeChar;
        byte uchLowByte = (byte)(uiUnicodeChar & 0x00FF);
        byte uchHighByte = (byte)((uiUnicodeChar >> 8) & 0x00FF);

        byte[] achArrayFor2StructuresOfTypeINPUT = null;
        int iSizeOfINPUTStructure = 0;

        if (Environment.Is64BitProcess)
        {
          // An array for pressing and unpressing of 'Н' button on a keyboard.
          byte[] achArray_ForChar_Russian_N_64bitWin =
          {
            // KEY_DOWN for the button with code v_uiVirtualKeyCode.
            // (0x00 + 0x04 = KEYEVENTF_UNICODE)
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 29, 4, 0x00 + 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,

            // KEY_UP for the button with code v_uiVirtualKeyCode.
            // (0x02 + 0x04 = KEYEVENTF_KEYUP | KEYEVENTF_UNICODE)
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 29, 4, 0x02 + 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
          };

          achArray_ForChar_Russian_N_64bitWin[10] = uchLowByte;
          achArray_ForChar_Russian_N_64bitWin[11] = uchHighByte;
          achArray_ForChar_Russian_N_64bitWin[50] = uchLowByte;
          achArray_ForChar_Russian_N_64bitWin[51] = uchHighByte;

          achArrayFor2StructuresOfTypeINPUT = achArray_ForChar_Russian_N_64bitWin;

          iSizeOfINPUTStructure = 40;
        }
        else
        {
          // An array for pressing and unpressing of 'Н' button on a keyboard.
          byte[] achArray_ForChar_Russian_N_32bitWin =
          {
            // KEY_DOWN for the button with code v_uiVirtualKeyCode.
            // (0x00 + 0x04 = KEYEVENTF_UNICODE)
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 29, 4, 0x00 + 0x04, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,

            // KEY_UP for the button with code v_uiVirtualKeyCode.
            // (0x02 + 0x04 = KEYEVENTF_KEYUP | KEYEVENTF_UNICODE)
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 29, 4, 0x02 + 0x04, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
          };

          achArray_ForChar_Russian_N_32bitWin[6] = uchLowByte;
          achArray_ForChar_Russian_N_32bitWin[7] = uchHighByte;
          achArray_ForChar_Russian_N_32bitWin[34] = uchLowByte;
          achArray_ForChar_Russian_N_32bitWin[35] = uchHighByte;

          achArrayFor2StructuresOfTypeINPUT = achArray_ForChar_Russian_N_32bitWin;

          iSizeOfINPUTStructure = 28;
        }

        UInt32 uiResultCode = SendInput(2, achArrayFor2StructuresOfTypeINPUT, iSizeOfINPUTStructure);

        if (uiResultCode != 2)
        {
          int iLastError = GetLastError();

          throw new Exception("Error in SendInput() function fom user32.dll. " +
            "GetLastError() = " + iLastError.ToString() + ", " +
            "uiUnicodeChar = '" + uiUnicodeChar + "', " +
            "v_sUnicodeString = \"" + v_sUnicodeString + "\".");
        }
      }
    }

    public static void Emulate_PressOf_Enter(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_RETURN);
    }

    public static void Emulate_PressOf_Tab(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_TAB);
    }

    public static void Emulate_PressOf_Space(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_SPACE);
    }

    public static void Emulate_PressOf_Left(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_LEFT);
    }

    public static void Emulate_PressOf_Right(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_RIGHT);
    }

    public static void Emulate_PressOf_Up(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_UP);
    }

    public static void Emulate_PressOf_Down(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_DOWN);
    }

    public static void Emulate_PressOf_Backspace(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_BACK);
    }

    public static void Emulate_PressOf_Del(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_DELETE);
    }

    public static void Emulate_PressOf_PageUp(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_PRIOR);
    }

    public static void Emulate_PressOf_PageDown(Process v_oTargetProcess)
    {
      Emulate_PressOf_Key_ByVirtualKeyCode(v_oTargetProcess, m_uiVirtualKey_for_VK_NEXT);
    }

    public static void Emulate_PressOf_Key_ByVirtualKeyCode(Process v_oTargetProcess, UInt16 v_uiVirtualKeyCode)
    {
      int iCurrentThreadId = GetCurrentThreadId();

      bool bFocusWindowFound = false;

      for (int iThreadIndex = 0; iThreadIndex < v_oTargetProcess.Threads.Count; iThreadIndex++)
      {
        ProcessThread oOtherThread = v_oTargetProcess.Threads[iThreadIndex];

        bool bAttachStatus = AttachThreadInput(
          iCurrentThreadId,
          oOtherThread.Id,
          true);

        if (!bAttachStatus)
        {
          int iLastError = GetLastError();

          throw new Exception(
            "Cannot attach to thread (id = " + oOtherThread.Id.ToString() + ") " +
            "for sending of emulated keyboard input.\n" +
            "GetLastError() == " + iLastError.ToString() + ",\n" +
            "v_uiVirtualKeyCode == " + v_uiVirtualKeyCode.ToString() + ".");
        }

        IntPtr uiFocusWindowHandle = GetFocus();
        if (uiFocusWindowHandle != (IntPtr)0)
        {
          bFocusWindowFound = true;

          EmulatePressOfKeyOnKeyboard_ByVirtualKeyCode(v_uiVirtualKeyCode);
        }

        bool bDetachStatus = AttachThreadInput(
          iCurrentThreadId,
          oOtherThread.Id,
          false);

        if (!bDetachStatus)
        {
          int iLastError = GetLastError();

          throw new Exception(
            "Cannot detach from thread (id = " + oOtherThread.Id.ToString() + ") " +
            "for sending of emulated keyboard input.\n" +
            "GetLastError() == " + iLastError.ToString() + ",\n" +
            "v_uiVirtualKeyCode == " + v_uiVirtualKeyCode.ToString() + ".");
        }

        if (bFocusWindowFound)
          break;
      }
    }

    public static void Emulate_TypingOfUnicodeString(Process v_oTargetProcess, String v_sUnicodeString)
    {
      int iCurrentThreadId = GetCurrentThreadId();

      bool bFocusWindowFound = false;

      for (int iThreadIndex = 0; iThreadIndex < v_oTargetProcess.Threads.Count; iThreadIndex++)
      {
        ProcessThread oOtherThread = v_oTargetProcess.Threads[iThreadIndex];

        bool bAttachStatus = AttachThreadInput(
          iCurrentThreadId,
          oOtherThread.Id,
          true);

        if (!bAttachStatus)
        {
          int iLastError = GetLastError();

          throw new Exception(
            "Cannot attach to thread (id = " + oOtherThread.Id.ToString() + ") " +
            "for sending of emulated keyboard input.\n" +
            "GetLastError() == " + iLastError.ToString() + ",\n" +
            "v_sUnicodeString == \"" + v_sUnicodeString + "\".");
        }

        IntPtr uiFocusWindowHandle = GetFocus();
        if (uiFocusWindowHandle != (IntPtr)0)
        {
          bFocusWindowFound = true;

          EmulatePressOfKeyOnKeyboard_ByUnicodeString(v_sUnicodeString);
        }

        bool bDetachStatus = AttachThreadInput(
          iCurrentThreadId,
          oOtherThread.Id,
          false);

        if (!bDetachStatus)
        {
          int iLastError = GetLastError();

          throw new Exception(
            "Cannot detach from thread (id = " + oOtherThread.Id.ToString() + ") " +
            "for sending of emulated keyboard input.\n" +
            "GetLastError() == " + iLastError.ToString() + ",\n" +
            "v_sUnicodeString == \"" + v_sUnicodeString + "\".");
        }

        if (bFocusWindowFound)
          break;
      }
    }

  }
}
