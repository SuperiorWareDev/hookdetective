Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Globalization

Public Class InvestigatorForm
    ' This will hold the last snapshot of memory we read
    'Private lastMemoryBlock() As Byte = Nothing
    ' P/Invoke declarations for memory reading
    Private Const PROCESS_VM_READ As Integer = &H10
    Private processHandle As IntPtr = IntPtr.Zero ' Our new handle variable!

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function OpenProcess(ByVal dwDesiredAccess As UInteger, <MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, ByVal dwProcessId As UInteger) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function ReadProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, ByVal lpBuffer As Byte(), ByVal nSize As UIntPtr, ByRef lpNumberOfBytesRead As UIntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CloseHandle(ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    ' This will hold the Process ID passed from our main form
    Private ReadOnly _targetProcessId As UInteger
    ' *** NEW: Variables to store our memory snapshots ***
    Private baselineMemoryBlock() As Byte = Nothing
    Private actionMemoryBlock() As Byte = Nothing
    ' A custom constructor that requires a Process ID to create this form
    ' *** UPGRADED: The constructor now accepts the Base Address! ***
    Public Sub New(targetProcessId As UInteger, baseAddress As IntPtr)
        InitializeComponent()

        _targetProcessId = targetProcessId
        Me.Text = $"Investigator - PID: {_targetProcessId}"

        ' *** NEW: Automatically fill the address box! ***
        ' Convert the address to a Hex string for the textbox
        If baseAddress <> IntPtr.Zero Then
            txtAddress.Text = baseAddress.ToString("X")
        End If
    End Sub

    Private Sub btnCaptureBaseline_Click(sender As Object, e As EventArgs) Handles btnCaptureBaseline.Click
        Dim startAddress As IntPtr
        Dim bytesToRead As Integer
        ' No longer need: Dim currentMemoryBlock() As Byte
        Dim bytesRead As UIntPtr = UIntPtr.Zero

        Try
            startAddress = New IntPtr(Long.Parse(txtAddress.Text, NumberStyles.HexNumber))
            bytesToRead = CInt(numSize.Value)
            ReDim baselineMemoryBlock(bytesToRead - 1) ' Use our class-level variable
        Catch ex As Exception
            MessageBox.Show("Invalid Start Address or Size.", "Input Error")
            Return
        End Try

        Dim tempProcessHandle As IntPtr = OpenProcess(PROCESS_VM_READ, False, _targetProcessId)
        If tempProcessHandle = IntPtr.Zero Then
            MessageBox.Show("Failed to open target process for Baseline scan.", "Process Error")
            Return
        End If

        Try
            ' Use baselineMemoryBlock directly
            If ReadProcessMemory(tempProcessHandle, startAddress, baselineMemoryBlock, New UIntPtr(CUInt(bytesToRead)), bytesRead) Then
                ' Adjust the size of baselineMemoryBlock if fewer bytes were read than requested
                If CUInt(bytesRead.ToUInt32()) < bytesToRead Then
                    ReDim Preserve baselineMemoryBlock(CInt(bytesRead.ToUInt32()) - 1)
                End If
                rtbBaselineDump.Text = FormatMemoryBlock(baselineMemoryBlock, startAddress, CUInt(bytesRead.ToUInt32()))
                rtbActionDump.Clear()
                MessageBox.Show("Baseline captured!", "Scan 1 Complete")
            Else
                rtbBaselineDump.Text = "Failed to read memory for Baseline."
                baselineMemoryBlock = Nothing ' Clear if failed
            End If
        Catch ex As Exception
            rtbBaselineDump.Text = $"Error reading memory: {ex.Message}"
            baselineMemoryBlock = Nothing ' Clear if failed
        Finally
            If tempProcessHandle <> IntPtr.Zero Then
                CloseHandle(tempProcessHandle)
            End If
        End Try
    End Sub
    Private Sub btnCaptureAction_Click(sender As Object, e As EventArgs) Handles btnCaptureAction.Click
        Dim startAddress As IntPtr
        Dim bytesToRead As Integer
        ' No longer need: Dim currentMemoryBlock() As Byte
        Dim bytesRead As UIntPtr = UIntPtr.Zero

        Try
            startAddress = New IntPtr(Long.Parse(txtAddress.Text, NumberStyles.HexNumber))
            bytesToRead = CInt(numSize.Value)
            ReDim actionMemoryBlock(bytesToRead - 1) ' Use our class-level variable
        Catch ex As Exception
            MessageBox.Show("Invalid Start Address or Size.", "Input Error")
            Return
        End Try

        Dim tempProcessHandle As IntPtr = OpenProcess(PROCESS_VM_READ, False, _targetProcessId)
        If tempProcessHandle = IntPtr.Zero Then
            MessageBox.Show("Failed to open target process for Action scan.", "Process Error")
            Return
        End If

        Try
            ' Use actionMemoryBlock directly
            If ReadProcessMemory(tempProcessHandle, startAddress, actionMemoryBlock, New UIntPtr(CUInt(bytesToRead)), bytesRead) Then
                ' Adjust the size of actionMemoryBlock if fewer bytes were read than requested
                If CUInt(bytesRead.ToUInt32()) < bytesToRead Then
                    ReDim Preserve actionMemoryBlock(CInt(bytesRead.ToUInt32()) - 1)
                End If
                rtbActionDump.Text = FormatMemoryBlock(actionMemoryBlock, startAddress, CUInt(bytesRead.ToUInt32()))
                MessageBox.Show("Action snapshot captured!", "Scan 2 Complete")
            Else
                rtbActionDump.Text = "Failed to read memory for Action."
                actionMemoryBlock = Nothing ' Clear if failed
            End If
        Catch ex As Exception
            rtbActionDump.Text = $"Error reading memory: {ex.Message}"
            actionMemoryBlock = Nothing ' Clear if failed
        Finally
            If tempProcessHandle <> IntPtr.Zero Then
                CloseHandle(tempProcessHandle)
            End If
        End Try
    End Sub
    Private Function FormatMemoryBlock(ByVal memoryBlock() As Byte, ByVal startAddress As IntPtr, ByVal bytesSuccessfullyRead As UInteger) As String
        Dim sb As New StringBuilder()
        For i = 0 To CInt(bytesSuccessfullyRead) - 1
            If i Mod 16 = 0 Then
                sb.Append($"{startAddress.ToInt64() + i:X16}: ")
            End If
            sb.Append($"{memoryBlock(i):X2} ")
            If (i + 1) Mod 16 = 0 AndAlso i > 0 Then
                sb.AppendLine()
            End If
        Next
        Return sb.ToString()
    End Function
    Private Function GenerateRtfWithComparison(ByVal displayBlock() As Byte, ByVal compareBlock() As Byte,
                                           ByVal startAddress As IntPtr, ByVal bytesSuccessfullyRead As UInteger,
                                           ByVal defaultColorIndex As Integer, ByVal changedColorIndex As Integer) As String
        Dim sb As New StringBuilder()
        ' RTF Header: Define the document, font, and a color table.
        ' Color 1=Gray (for addresses), Color 2=LimeGreen, Color 3=White (default text), Color 4=Red
        sb.Append("{\rtf1\ansi\deff0{\fonttbl{\f0 Courier New;}}")
        sb.Append("{\colortbl;\red128\green128\blue128;\red0\green255\blue0;\red255\green255\blue255;\red255\green0\blue0;}") ' Added Red as color 4
        sb.Append("\fs20 ") ' Set font size

        For i = 0 To CInt(bytesSuccessfullyRead) - 1
            If i Mod 16 = 0 Then
                sb.Append($"\par\cf1 {startAddress.ToInt64() + i:X16}: ") ' \par=new line, \cf1=color 1 (gray)
            End If

            Dim useColorIndex As Integer = defaultColorIndex
            ' Check if this byte has changed compared to the other block
            If compareBlock IsNot Nothing AndAlso i < compareBlock.Length AndAlso displayBlock(i) <> compareBlock(i) Then
                useColorIndex = changedColorIndex ' If it changed, use the specified changed color
            End If

            sb.Append($"\cf{useColorIndex} {displayBlock(i):X2} ") ' Append the byte with its chosen color
        Next

        sb.Append("}") ' Close the RTF document
        Return sb.ToString()
    End Function
    Private Sub InvestigatorForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnCompareDumps_Click(sender As Object, e As EventArgs) Handles btnCompareDumps.Click
        If baselineMemoryBlock Is Nothing OrElse actionMemoryBlock Is Nothing Then
            MessageBox.Show("Please capture both a Baseline and an Action snapshot first.", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Ensure addresses and sizes match for a fair comparison, or compare up to the shorter length
        Dim compareLength As Integer = Math.Min(baselineMemoryBlock.Length, actionMemoryBlock.Length)
        Dim startAddress As IntPtr
        Try
            startAddress = New IntPtr(Long.Parse(txtAddress.Text, NumberStyles.HexNumber))
        Catch ex As Exception
            MessageBox.Show("Invalid Start Address.", "Input Error")
            Return
        End Try

        ' Update rtbBaselineDump: default White, changed bytes Red
        rtbBaselineDump.Rtf = GenerateRtfWithComparison(baselineMemoryBlock, actionMemoryBlock, startAddress, CUInt(compareLength), 3, 4)

        ' Update rtbActionDump: default White, changed bytes Green
        rtbActionDump.Rtf = GenerateRtfWithComparison(actionMemoryBlock, baselineMemoryBlock, startAddress, CUInt(compareLength), 3, 2)

        MessageBox.Show("Comparison complete. Differences highlighted (Red in Baseline, Green in Action).", "Analysis Done")
    End Sub
End Class