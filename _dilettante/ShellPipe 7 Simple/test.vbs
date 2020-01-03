Option Explicit
'
'Sample script that is run by the
'ShellPipe demonstration program.
'

Dim dicStrings, I

'Get input to process.
Set dicStrings = CreateObject("Scripting.Dictionary")
With WScript.StdIn
    I = 1
    Do Until .atEndOfStream
        dicStrings.Add I, .ReadLine()
        I = I + 1
    Loop
End With

'Process and return results.
With WScript.StdOut
    For I = 1 To dicStrings.Count
        WScript.StdErr.WriteLine dicStrings.Item(I)
        .WriteLine StrReverse(dicStrings.Item(I))
    Next
    dicStrings.RemoveAll

    'Example of dangling output (no NewLine).
    .Write "**Finished**"

    'Example of EOF (linger so EOF is detected before we end).
    .Close
    WScript.Sleep 500
End With

WScript.Quit I - 1
