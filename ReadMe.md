VBA-challenge assingment - Module 2 Homework
Includes 3 screenshots(1 per worksheet of the top of each worksheet), the final excel file, the separate extracted script files, and this README.

The code: 'Looping through worksheets
Dim ws As Excel.Worksheet
    For Each ws In Worksheets
        ws.Activate
    Next
was found in collaboration with coursemate Patricia Ferreira as well as this youtube video: https://www.youtube.com/watch?v=3OfVIsKy59c&t=174s

The code: 'Setting up last row variable for summarized data specifically
Dim lastRow As Long
lastRow = Range("I1").End(xlDown).Row
was used in this assignment with information from this website: https://excelchamps.com/vba/find-last-row-column-cell/