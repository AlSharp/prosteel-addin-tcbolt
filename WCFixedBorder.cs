// Name: WCFixedBorder.cs 
// 
// Description: 
// Class helps us to set border of WindowContent's Zone to fixed. 
// It uses Win32 API to change form's style flags. If form is not 
// found (when zone is docked), it waits for first undocking and 
// then sets forms style. 
// 
// Author: Dušan Paulovič (GISoft) www.gisoft.cz
// 
// Usage: 
// WindowContent WC = WindowManager.GetForMicroStation().DockPanel(...); 
// AdapterWorkarounds.WCFixedBorder.SetBorderFixed(WC); 
//  
// Tested on: MicroStation V8i, version: 08.11.07.74 
// 
// Warranty: 
// THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED  
// BY APPLICABLE LAW. EXCEPT WHEN OTHERWISE STATED IN WRITING THE  
// COPYRIGHT HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM "AS IS"  
// WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  
// INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF  
// MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.  
// THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM  
// IS WITH YOU. SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE  
// COST OF ALL NECESSARY SERVICING, REPAIR OR CORRECTION. 

using System; 
using System.Collections.Generic; 
using System.Runtime.InteropServices; 
using Bentley.Windowing; 
using Bentley.Windowing.Docking; 
using System.Windows.Forms; 

namespace AdapterWorkarounds 
{ 
  public class WCFixedBorder 
  { 
    [DllImport("user32.dll")] 
    private static extern IntPtr GetParent(IntPtr hWnd); 

    [DllImport("user32.dll", SetLastError = true)] 
    private static extern int  
      GetWindowLong(IntPtr hWnd, int nIndex); 

    [DllImport("user32.dll")] 
    private static extern int  
      SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong); 

    [DllImport("user32.dll")] 
    [return: MarshalAs(UnmanagedType.Bool)] 
    private static extern bool  
      SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,  
      int X, int Y, int cx, int cy, uint uFlags); 

    private static Dictionary<Zone, List<WindowContent>>  
      ZoneContentsList = new Dictionary<Zone, List<WindowContent>>();

    [Flags] 
    private enum SWP 
    { 
      SWP_NOSIZE = 0x0001, 
      SWP_NOMOVE = 0x0002, 
      SWP_NOZORDER = 0x0004, 
      SWP_NOREDRAW = 0x0008, 
      SWP_NOACTIVATE = 0x0010, 
      SWP_FRAMECHANGED = 0x0020, 
      SWP_SHOWWINDOW = 0x0040, 
      SWP_HIDEWINDOW = 0x0080, 
      SWP_NOCOPYBITS = 0x0100, 
      SWP_NOOWNERZORDER = 0x0200, 
      SWP_NOSENDCHANGING = 0x0400, 
      SWP_DRAWFRAME = SWP_FRAMECHANGED, 
      SWP_NOREPOSITION = SWP_NOOWNERZORDER, 
      SWP_DEFERERASE = 0x2000, 
      SWP_ASYNCWINDOWPOS = 0x4000 
    } 

    // base method, sets size of the form, a border of the form; 
    // disables minimize and maximize boxes  
    private static bool _setBorderFixed(Zone z) 
    { 
      Form frm = z.FindForm(); 
      if (frm != null) 
      { 
        const int WS_THICKFRAME = 0x40000;
        const int WS_MINIMIZEBOX = 0x20000;
        const int WS_MAXIMIZEBOX = 0x10000;
        const int GWL_STYLE = -16;

        SetWindowLong(frm.Handle, GWL_STYLE, GetWindowLong(frm.Handle, GWL_STYLE) & ~WS_THICKFRAME);
        SetWindowLong(frm.Handle, GWL_STYLE, GetWindowLong(frm.Handle, GWL_STYLE) & ~WS_MAXIMIZEBOX);
        SetWindowLong(frm.Handle, GWL_STYLE, GetWindowLong(frm.Handle, GWL_STYLE) & ~WS_MINIMIZEBOX);

        SetWindowPos(frm.Handle, new IntPtr(1), 0, 0, 283, 254,  
          (uint)(SWP.SWP_SHOWWINDOW |  
          SWP.SWP_FRAMECHANGED | SWP.SWP_NOMOVE | SWP.SWP_NOZORDER)); 

        return true; 
      } 
      return false; 
    } 

    private static EventHandler  
      ParentChangedEvent = new EventHandler(Zone_ParentChanged); 
    public static void SetBorderFixed(WindowContent content) 
    { 
      if (!_setBorderFixed(content.Zone)) 
      { 
        if (content.Zone != null) 
        { 
          content.Zone.ParentChanged += ParentChangedEvent; 
          if (ZoneContentsList.ContainsKey(content.Zone)) 
            ZoneContentsList[content.Zone].Add(content); 
          else 
          { 
            List<WindowContent> wcList = new List<WindowContent>(); 
            wcList.Add(content); 
            ZoneContentsList.Add(content.Zone, wcList); 
          } 
        } 
      } 
    } 
     
    // If form was not found, it will be handled now... 
    private static void Zone_ParentChanged(object sender, EventArgs e) 
    { 
      if (sender is Zone) 
      { 
        Zone zone = sender as Zone; 
        if (ZoneContentsList.ContainsKey(zone)) 
        { 
          for (int id = ZoneContentsList[zone].Count - 1; id >= 0; --id) 
          { 
            WindowContent wc = ZoneContentsList[zone][id]; 
            if (_setBorderFixed(wc.Zone)) 
            { 
              ZoneContentsList[zone].Remove(wc); 
            } 
          } 

          if (ZoneContentsList[zone].Count == 0) 
          { 
            ZoneContentsList.Remove(zone); 
            zone.ParentChanged -= ParentChangedEvent; 
          } 
        } 
      } 
    } 
  } 
}