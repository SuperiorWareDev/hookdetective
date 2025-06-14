# HookDetective v1 - Public Demonstration

## Overview

HookDetective is a tool designed for real-time analysis of Windows applications. It allows users to inspect window properties, extract process information, and even interact with running processes. This version demonstrates the core functionality of the tool.

## Key Features

* **Real-Time Window Inspection:** Displays detailed information about any window under the cursor, including its handle, class name, and process ID.
* **Process Information Extraction:** Retrieves process name, path, and other relevant details.
* **Target Finder:** Allows users to precisely select a window for analysis.
* **Freeze Mode:** Pauses the real-time analysis for detailed inspection.

## Technology Stack

This project leverages a hybrid approach to achieve both rapid development and maximum performance:

* **VB.NET:** The core application logic and user interface are built in VB.NET for its ease of use and rapid development capabilities.
* **C++ & Assembly:** Performance-critical components, such as the hook procedures and low-level memory access, are implemented in C++ and Assembly for optimal efficiency.

## Version Notes

This is a demonstration version of HookDetective. The complete version includes advanced features, such as:

* Custom `.dll` drivers for enhanced system-level access.
* Proprietary algorithms for advanced memory analysis.

These features are not included in this public release for security and intellectual property reasons.

## Usage

1.  **Download:** Clone this repository to your local machine.
2.  **Build:** Open the `HookDetective.sln` solution file in Visual Studio.
3.  **Run:** Build and run the `HookDetective` project.

Once the application is running, you can:

* Hover your mouse over any window to see its properties in real-time.
* Use the "Target Finder" button to select a specific window.
* Press F2 to toggle "Freeze Mode" and pause the analysis.

## Contributing

This is not an open-source project. Contributions are not accepted.

## License

This project is for demonstration purposes only. All rights reserved.
