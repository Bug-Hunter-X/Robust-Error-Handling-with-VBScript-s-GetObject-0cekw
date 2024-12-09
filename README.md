# Robust Error Handling with VBScript's GetObject

This example demonstrates how to handle errors gracefully when using VBScript's `GetObject` function. The `GetObject` function can fail for various reasons, such as the specified file not existing or being inaccessible, or problems with the COM object itself.

The original code lacks proper error handling, which can lead to script crashes or unexpected behavior. The solution shows how to implement robust error handling to prevent these issues. 

## Bug
The `bug.vbs` file contains code that uses `GetObject` without proper error handling. If the specified file is not found, the script will crash.

## Solution
The `bugSolution.vbs` file shows the improved code that includes comprehensive error handling using `On Error Resume Next` and explicit error checking.