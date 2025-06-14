Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1
    ' (All the other GetWindowText, GetClassName, etc. declarations remain)
    ' (Your other variable declarations are here)
    Private isFrozen As Boolean = False ' Our on/off switch for freeze mode!
    ' --- Keep track of the last window we've targeted ---
    Private lastWindowHighlighted As IntPtr = IntPtr.Zero

    ' (All the other GetWindowText, GetClassName, etc. declarations remain)
    ' *** UPDATED: P/Invoke with the new GetParent function ***
    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetParent(ByVal hWnd As IntPtr) As IntPtr
    End Function

    ' (All other P/Invoke declarations remain the same as before)
    Private Delegate Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Private hookProc As HookCallback
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As HookCallback, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Private Shared Function WindowFromPoint(ByVal p As POINT) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetClassName(ByVal hWnd As IntPtr, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private hHook As IntPtr = IntPtr.Zero
    Private highlighter As New HighLightForm()
    Private lastHwnd As IntPtr = IntPtr.Zero

    Private isExiting As Boolean = False ' NEW: To allow a real exit

    ' (Structures remain the same)
    Public Structure RECT
        Public Left As Integer, Top As Integer, Right As Integer, Bottom As Integer
    End Structure
    Private Structure POINT
        Public x As Integer, y As Integer
    End Structure
    Private Structure MOUSEHOOKSTRUCT
        Public pt As POINT, hwnd As IntPtr, wHitTestCode As UInteger, dwExtraInfo As IntPtr
    End Structure

    ' *** MAJOR UPDATE: Hook callback now populates our WindowDetails object ***
    Private Function MouseHookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If isFrozen Then Return CallNextHookEx(hHook, nCode, wParam, lParam)

        If nCode >= 0 AndAlso wParam = WM_MOUSEMOVE Then
            Dim hookStruct As MOUSEHOOKSTRUCT = CType(Marshal.PtrToStructure(lParam, GetType(MOUSEHOOKSTRUCT)), MOUSEHOOKSTRUCT)
            Dim hWnd As IntPtr = WindowFromPoint(hookStruct.pt)

            If hWnd <> lastHwnd Then
                lastHwnd = hWnd
                If hWnd <> IntPtr.Zero Then
                    Dim details As New WindowDetails()
                    Dim pid As UInteger = 0
                    GetWindowThreadProcessId(hWnd, pid)

                    Dim sbCaption As New StringBuilder(256)
                    GetWindowText(hWnd, sbCaption, sbCaption.Capacity)
                    Dim sbClass As New StringBuilder(256)
                    GetClassName(hWnd, sbClass, sbClass.Capacity)

                    ' Populate our new details object
                    details.Handle = hWnd.ToString("X") ' Display as Hex
                    details.Caption = sbCaption.ToString()
                    details.ClassName = sbClass.ToString()
                    details.ProcessID = pid
                    details.ParentHandle = GetParent(hWnd).ToString("X")
                    'details.BaseAddress = proc.MainModule.BaseAddress
                    Try ' Getting process path can fail on system processes
                        details.ProcessPath = Process.GetProcessById(CInt(pid)).MainModule.FileName
                    Catch ex As Exception
                        details.ProcessPath = "Access Denied"
                    End Try

                    ' Update the PropertyGrid on the UI thread
                    Me.BeginInvoke(Sub() PropertyGrid1.SelectedObject = details)

                    ' Update highlighter
                    Dim windowRect As New RECT()
                    GetWindowRect(hWnd, windowRect)
                    highlighter.SetBounds(windowRect.Left, windowRect.Top, windowRect.Right - windowRect.Left, windowRect.Bottom - windowRect.Top)
                    highlighter.Show()
                Else
                    highlighter.Hide()
                    Me.BeginInvoke(Sub() PropertyGrid1.SelectedObject = Nothing)
                End If
            End If
        End If
        Return CallNextHookEx(hHook, nCode, wParam, lParam)
    End Function

    ' *** UPDATED: Form Load and Closing Events for System Tray ***
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        hookProc = AddressOf MouseHookCallback
        hHook = SetWindowsHookEx(WH_MOUSE_LL, hookProc, GetModuleHandle(Nothing), 0)

        ' Setup the NotifyIcon
        NotifyIcon1.Icon = Me.Icon ' Use the form's own icon
        NotifyIcon1.Visible = True
        ' Set the form's opacity to 50% (semi-transparent)
        Me.Opacity = 0.4
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ' Check if the key pressed was F2
        If e.KeyCode = Keys.F2 Then
            ' Flip our frozen state (true to false, false to true)
            isFrozen = Not isFrozen

            ' Give some visual feedback by changing the window title
            If isFrozen Then
                Me.Text = "The Watcher (FROZEN)"
            Else
                Me.Text = "The Watcher"
            End If
        End If
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        ' If we ARE exiting, clean up everything
        UnhookWindowsHookEx(hHook)
        NotifyIcon1.Dispose()

    End Sub
    ' --- This is the MAIN capture logic. It gets all details for a window handle ---
    Private Sub CaptureWindowInfo(ByVal hWnd As IntPtr)
        If hWnd = IntPtr.Zero Then
            PropertyGrid1.SelectedObject = Nothing
            Return
        End If

        Dim details As New WindowDetails()
        Dim pid As UInteger = 0
        GetWindowThreadProcessId(hWnd, pid)

        Dim sbCaption As New StringBuilder(256)
        GetWindowText(hWnd, sbCaption, sbCaption.Capacity)

        Dim sbClass As New StringBuilder(256)
        GetClassName(hWnd, sbClass, sbClass.Capacity)

        details.Handle = hWnd.ToString("X")
        details.Caption = sbCaption.ToString()
        details.ClassName = sbClass.ToString()
        details.ProcessID = pid
        details.ParentHandle = GetParent(hWnd).ToString("X")


        Try
            Dim proc As Process = Process.GetProcessById(CInt(pid))

            ' *** NEW DIAGNOSTIC LINE: Let's see the process name! ***
            Me.Text = $"The Watcher - Found: {proc.ProcessName}"

            details.ProcessPath = proc.MainModule.FileName
            details.BaseAddress = proc.MainModule.BaseAddress
        Catch ex As Exception
            ' We can leave the MessageBox here just in case, but it's not firing.
            MessageBox.Show($"I failed. The secret error is: {ex.Message}", "Diagnostic Message")
            details.ProcessPath = "Access Denied"
            details.BaseAddress = IntPtr.Zero
        End Try

        PropertyGrid1.SelectedObject = details
    End Sub
    ' --- Event handlers for our new Target Finder button ---



    Private Sub btnTargetFinder_MouseUp(sender As Object, e As MouseEventArgs) Handles btnTargetFinder.MouseUp
        ' When we release the mouse button, our work is done

        ' Erase the final highlight rectangle
        DrawHighlightRectangle()

        ' Restore the default cursor
        Me.Cursor = Cursors.Default

        ' This is it! Capture the final window's information and display it
        CaptureWindowInfo(lastWindowHighlighted)
        lastWindowHighlighted = IntPtr.Zero
    End Sub

    ' --- Helper function to draw a temporary highlight rectangle ---
    Private Sub DrawHighlightRectangle()
        If lastWindowHighlighted = IntPtr.Zero Then Return

        Dim rect As RECT
        GetWindowRect(lastWindowHighlighted, rect)
        ControlPaint.DrawReversibleFrame(New Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top), Color.Black, FrameStyle.Thick)
    End Sub


    ' --- Event Handler for our new "Launch Investigator" button ---
    Private Sub btnLaunchInvestigator_Click(sender As Object, e As EventArgs) Handles btnLaunchInvestigator.Click



        ' This logic is the same as your old double-click event [cite: 1]
        If PropertyGrid1.SelectedObject Is Nothing Then Return

        Dim details As WindowDetails = TryCast(PropertyGrid1.SelectedObject, WindowDetails)

        If details IsNot Nothing AndAlso details.ProcessID > 0 Then
            ' *** UPGRADED: Pass the BaseAddress along with the PID ***
            Dim investigator As New InvestigatorForm(details.ProcessID, details.BaseAddress)
            investigator.Show()
        End If
    End Sub
    ' *** NEW: Event Handlers for our Context Menu Items ***
    Private Sub ShowHideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowHideToolStripMenuItem.Click
        Me.Visible = Not Me.Visible
    End Sub

    Private Sub FreezeSpyingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FreezeSpyingToolStripMenuItem.Click
        isFrozen = Not isFrozen
        If isFrozen Then
            FreezeSpyingToolStripMenuItem.Text = "Unfreeze Spying"
            highlighter.Hide() ' Hide highlighter when frozen
        Else
            FreezeSpyingToolStripMenuItem.Text = "Freeze Spying"
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        isExiting = True ' Set the flag to allow a true exit
        Application.Exit()
    End Sub

    ' *** NEW: Double click the tray icon to show/hide the form ***
    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        ShowHideToolStripMenuItem.PerformClick()
    End Sub

    Private Sub PropertyGrid1_Click(sender As Object, e As EventArgs) Handles PropertyGrid1.Click

    End Sub

    Private Sub ContextMenuStrip1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ContextMenuStrip1.MouseDoubleClick

    End Sub

    Private Sub btnTargetFinder_Click(sender As Object, e As EventArgs) Handles btnTargetFinder.Click

    End Sub
    Private Sub btnTargetFinder_MouseDown(sender As Object, e As MouseEventArgs) Handles btnTargetFinder.MouseDown
        If isFrozen Then Return ' If we are frozen, do nothing!

        ' When the mouse button is pressed, change the cursor to a crosshair
        Me.Cursor = Cursors.Cross
    End Sub
    Private Sub btnTargetFinder_MouseMove(sender As Object, e As MouseEventArgs) Handles btnTargetFinder.MouseMove
        If isFrozen Then Return ' If we are frozen, do nothing!

        ' Only run this if we are holding the left mouse button down
        If e.Button <> MouseButtons.Left Then Return



        ' Get the handle of the window currently under the cursor
        ' Create a new instance of OUR custom POINT structure
        Dim cursorPoint As New POINT()
        ' Manually copy the X and Y values from the cursor's position
        cursorPoint.x = Cursor.Position.X
        cursorPoint.y = Cursor.Position.Y

        ' Now, pass OUR custom point structure to the API function
        Dim currentWindow As IntPtr = WindowFromPoint(cursorPoint)

        ' If the window is new, update the highlight
        If currentWindow <> lastWindowHighlighted Then
            ' Erase the old rectangle
            DrawHighlightRectangle()

            ' Set the new window handle and draw the new rectangle
            lastWindowHighlighted = currentWindow
            DrawHighlightRectangle()
        End If




    End Sub

    Private Sub btnFindByName_Click(sender As Object, e As EventArgs) Handles btnFindByName.Click
        ' First, make sure we have spied on a window and have details in the grid.
        If PropertyGrid1.SelectedObject Is Nothing Then
            MessageBox.Show("Please use the Target Finder 🎯 to select a window first.", "No Target", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim details As WindowDetails = TryCast(PropertyGrid1.SelectedObject, WindowDetails)
        If details Is Nothing OrElse String.IsNullOrEmpty(details.ProcessPath) Then
            MessageBox.Show("Could not get process details from the selected window.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Try
            ' --- This is our new logic! ---

            ' 1. Get just the file name (e.g., "SamClient11") from the full path, without the ".exe"
            Dim processName As String = System.IO.Path.GetFileNameWithoutExtension(details.ProcessPath)

            ' 2. Search all running processes for that specific name.
            Dim processes() As Process = Process.GetProcessesByName(processName)

            ' 3. Check if we found it.
            If processes.Length = 0 Then
                MessageBox.Show($"Could not find any running process named '{processName}'.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Return
            End If

            ' 4. Get the first process we found (usually there's only one).
            Dim targetProcess As Process = processes(0)

            ' 5. Go directly for the Base Address!
            Dim baseAddress As IntPtr = targetProcess.MainModule.BaseAddress

            ' --- Show the result immediately! ---
            MessageBox.Show($"Success! Found Base Address by name: {baseAddress.ToString("X")}", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Also update our details object and the grid view
            details.BaseAddress = baseAddress
            PropertyGrid1.SelectedObject = Nothing ' A little trick to force the grid to refresh
            PropertyGrid1.SelectedObject = details

        Catch ex As Exception
            ' If this fails, we will know the EXACT reason.
            MessageBox.Show($"Failed to get Base Address by name. The error is: {ex.Message}", "Direct Access Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' This logic is the same as your old double-click event [cite: 1]
        If PropertyGrid1.SelectedObject Is Nothing Then Return

        'Dim details As WindowDetails = TryCast(PropertyGrid1.SelectedObject, WindowDetails)

        If details IsNot Nothing AndAlso details.ProcessID > 0 Then
            ' *** UPGRADED: Pass the BaseAddress along with the PID ***
            Dim investigator As New InvestigatorForm(details.ProcessID, details.BaseAddress)
            investigator.Show()
 
        End If

    End Sub

End Class
' *** NEW: A class to hold all our window information for the PropertyGrid ***
Public Class WindowDetails
    <Category("Identity"), DisplayName("Window Handle (hWnd)")>
    Public Property Handle As String

    <Category("Identity"), DisplayName("Window Caption")>
    Public Property Caption As String

    <Category("Identity"), DisplayName("Class Name")>
    Public Property ClassName As String

    <Category("Process"), DisplayName("Process ID (PID)")>
    Public Property ProcessID As UInteger

    <Category("Process"), DisplayName("Program Path")>
    Public Property ProcessPath As String

    <Category("Hierarchy"), DisplayName("Parent Handle (hWnd)")>
    Public Property ParentHandle As String
    <Category("Process"), DisplayName("Base Address")>
    Public Property BaseAddress As IntPtr
End Class