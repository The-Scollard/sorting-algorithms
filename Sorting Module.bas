Attribute VB_Name = "modSorting"
'Sorting Algorithms
'Luke Scollard
'ICS 3U1
'05/16/14

'Software Definition:
'The purpose of this code is to read lists of numbers, then sort them into a list that goes from least to greatest,
'then output the list into a text file. The program reads the file using batch processing, and records the time it
'takes for the sort to run. The program also checks to see whether the outputed list is actually sorted, alerting the
'user if the list is not sorted.

'Design Decisions:
'Some design decisions include a subroutine to check whether or not the list is actually sorted. This allows the user
'to see whether or not the sort is actually working. This allows the user to be alerted to any problems with the code
'without having to go into the sorted lists to check.

'The program uses a batch processing system to sort through files. All of the file names are contained within another
'file, which the program reads to find the files it needs to sort. This allows the program to sort through all the files
'next to no input from the user, allowing the program to sort faster.

'Another design decision is to use Heap Sort as the fifth sort in the algorithm. This sort was chosen due to it being
'easier to implement and understand than Quick Sort.

'A design decision was that the program runs in a module, and sorts the files with next to no input from the user.

'A constraint of the program is that it can only sort lists of numbers from the Data section of the Grade 11 Computer
'Science folder. If a list is located in any other directory, the program will not be able to sort it. This is because
'the file reads the ending of the file, which is then concotinated to the directory that all the files share, to allow
'the program to find the file.

'A design decision was to only go up to 50000 elements for Bubble Sort, Selection Sort, and Insertion Sort, and 1000000
'elements for Heap Sort and Shell Sort. This was due to the long amounts of time it would take for the first three sorts
'go through and sort the longer lists. Shell Sort and Heap Sort, due to how fast they are, could sort through the longer
'lists in comparitively short amounts of time, so that's why they sort up to 1000000 elements.

'Variable Dictionary

'Variable Name         Scope           Type         Purpose
'num()                 Global          Long         This variable is used to store the numbers from the list the
'                                                   the program is reading. It stores the number that is found at
'                                                   a certain position in the array.
'Location              Global          String       This varible stores the name of the directory where the file names
'                                                   are stored. It is concotinated onto the front of the file names to
'                                                   allow the program to find the specific file.
'ele                   Global          Long         This variable counts the number of elements found in the list. It
'                                                   is also used to dimension the num() array so that the array is the
'                                                   exact size of the list being read.
'Sort_Time             Global          Single       This variable contains the amount of time the sort took to sort all
'                                                   the data from the list into it's proper order.
'Sort_Type             Global          String       This variable is recorded from the file, and it tells the program which
'                                                   specific sort to use to sort the data. The program checks what the variable
'                                                   contains to see which sort to use.
'File_Name             Global          String       This variable is recorded from the file, and it is concatinated to the end
'                                                   of the variable Location to give the full file location. The variable allows
'                                                   the program to find the specific file to read and sort.
                                                         
                                                         

Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Global num() As Long
Global Location As String
Global ele As Long
Global Sort_Time As Single
Global Sort_Type As String
Global File_Name As String

Sub Main()

'Purpose: The purpose of this subroutine is to go through the file "Sorting Data" and read the file names from there. It
'activates the various other subroutines that find the file, sort the file, output the file, and check the file.
 
'Variable Dictionary

'Variable Name         Scope           Type         Purpose
'Start_Time            Sub Main()      Single       This variable records the time at which the sort is started. It is subtracted
'                                                   from the end time to find the amount of time it takes for the sort to sort through
'                                                   the file.
'End_Time              Sub Main()      Single       This variable records the time at which the sort ends. The start time is subtracted
'                                                   from it to find the amount of time it takes for the sort to sort the file.

    Dim Start_Time As Single
    Dim End_Time As Single
    Location = "\\SS17\S260 Classes$\G11 Computer Science\Assignments\Data\"
    
    Open "U:\Documents\Sorting Data.txt" For Input As #1
        Do Until EOF(1) = True
            Input #1, File_Name, Sort_Type
            Call Find_File
            Start_Time = GetTickCount()
            Call Sort_File
            End_Time = GetTickCount()
            Sort_Time = (End_Time - Start_Time) / 1000
            Call Write_File
            Call Check_File
        Loop
    Close #1
      
End Sub

Sub Find_File()

'Purpose: The purpose of this subroutine is to read the file, and copy the values into the num array.
    
    ReDim num(0)
    ele = 0
    
    Open Location & File_Name For Input As #2
        Do Until EOF(2) = True
            ele = ele + 1
            ReDim Preserve num(ele)
            Input #2, num(ele)
        Loop
    Close #2
    
End Sub

Sub Sort_File()

'Purpose: The purpose of this subroutine is to figure out from the data in the file "Sorting Data" the sort that will
'be used to sort the data.
    
    If Sort_Type = "Bubble Sort" Then
        Call Bubble_Sort
    ElseIf Sort_Type = "Selection Sort" Then
        Call Selection_Sort
    ElseIf Sort_Type = "Insertion Sort" Then
        Call Insertion_Sort
    ElseIf Sort_Type = "Shell Sort" Then
        Call Shell_Sort
    ElseIf Sort_Type = "Heap Sort" Then
        Heap_Sort.Heap_Sort
    End If
    
End Sub

Sub Bubble_Sort()

'Purpose: The purpose of this subroutine is to sort the data using the Bubble Sort method.

'Variable Dictionary

'Variable Name         Scope           Type         Purpose
'k                     Bubble_Sort()   Long         This variable acts as the counter for a For Loop in the Bubble Sort.
'N1                    Bubble_Sort()   Long         This variable is used to record the value of a number when switching
'                                                   the place of 2 numbers during the sort.
'N2                    Bubble_Sort()   Long         This variable is used to record the value of the other number when
'                                                   switching the place of 2 numbers during the sort.
'Switched              Bubble_Sort()   Boolean      The boolean checks to see if any switches were made during the run
'                                                   through of the list.
    Dim k As Long
    Dim N1 As Long
    Dim N2 As Long
    Dim Switched As Boolean
    Switched = True
    Do Until Switched = False 'If nothing is switched, then the list is sorted and Bubble Sort ends
        Switched = False
        For k = 0 To (ele - 2)
            If num(ele - k) < num(ele - k - 1) Then 'If the number to the left in the list is greater, then numbers are switched.
                N1 = num(ele - k)
                N2 = num(ele - k - 1)
                num(ele - k) = N2
                num(ele - k - 1) = N1
                Switched = True
            End If
        Next k
    Loop
End Sub

Sub Write_File()

'Purpose: The purpose of this subroutine is to write the sorted data into a new file, and write the sort time into the appropriate
'file.

'Variable Dictionary

'Variable Name         Scope           Type         Purpose
'c                     Write_File()    Long         This variable is used as a counter in a For Loop.
    
    Dim c As Long
    
    Open "U:\Documents\" & Sort_Type & "-" & File_Name For Output As #3
        For c = 1 To ele
            Write #3, num(c)
        Next c
    Close #3
    
    If Left$(File_Name, 6) = "Random" Then
        Open "U:\Documents\" & "Random" & "-" & Sort_Type & ".txt" For Append As #4
            Write #4, ele, Sort_Time
        Close #4
    ElseIf Left$(File_Name, 8) = "Reversed" Then
        Open "U:\Documents\" & "Reversed" & "-" & Sort_Type & ".txt" For Append As #5
            Write #5, ele, Sort_Time
        Close #5
    ElseIf Left$(File_Name, 9) = "Ascending" Then
        Open "U:\Documents\" & "Ascending" & "-" & Sort_Type & ".txt" For Append As #6
            Write #6, ele, Sort_Time
        Close #6
    End If
            
End Sub

Sub Check_File()

'Purpose: The purpose of this subroutine is to check to see if the file is properly sorted, and to alert the user if
'there any of the lists were not properly sorted by the program.

'Variable Dictionary

'Variable Name         Scope           Type         Purpose
'Sorted                Check_File()    Boolean      This boolean is used to check if the file is sorted. If any number is
'                                                   found out of place, the boolean is changed to indicate the list isn't
'                                                   sorted.
'r                     Check_File()    Long         This variable is used as a counter in a For Loop.

    Dim Sorted As Boolean
    Dim r As Long
    
    Sorted = True
    
    For r = 1 To (ele - 1)
        If num(r) > num(r + 1) Then   'If the number after the checked number is greater than the checked number, the file
            Sorted = False            'is marked as not sorted.
        End If
    Next r
    
    If Sorted = False Then
        MsgBox (File_Name & "-" & Sort_Type & " is not properly sorted") 'If the file isn't sorted, the user is alerted.
    End If
End Sub

Sub Selection_Sort()

'Purpose: The purpose of this subroutine is to sort the data using the Selection Sort method.

'Variable Dictionary

'Variable Name         Scope             Type       Purpose
'Sml                   Selection_Sort()  Long       This variable holds the value of the smallest number found in the
'                                                   current run-through of the list.
'N                     Selection_Sort()  Long       This variable holds the value of the number that is found at the
'                                                   beginning of the run-through, so that it can be deposited back into
'                                                   the list after the smallest number is found.
'P                     Selection_Sort()  Long       This variable records the position of the number found at the
'                                                   beginning of the run-through, so that is can be deposited back into
'                                                   the list after the smallest number is found.
'c                     Selection_Sort()  Long       This variable is used as a counter in a For Loop.
'r                     Selection_Sort()  Long       This variable is used as a counter in a For Loop

    Dim Sml As Long
    Dim N As Long
    Dim P As Long
    Dim c As Long
    Dim r As Long
    
    For c = 1 To ele
        Sml = num(c)
        For r = c To ele
            If num(r) < Sml Then 'If the checked number is less then the smallest number, the smallest number becomes the checked number.
                Sml = num(r)
                N = num(c)
                P = r
            End If
        Next r
        num(P) = N
        num(c) = Sml
    Next c
End Sub

Sub Insertion_Sort()

'Purpose: The purpose of this subroutine is to sort the data using the Insertion Sort method.

'Variable Dictionary

'Variable Name         Scope             Type       Purpose
'r                     Insertion_Sort()  Long       This variable is used as a counter in a For Loop.
'c                     Insertion_Sort()  Long       This variable is used as a counter in a For Loop.
'Compare               Insertion_Sort()  Long       This variable takes the value of the number that will be compared in the
'                                                   Insertion Sort, so that it can be compared without the number itself being
'                                                   moved, until it needs to be.

    Dim r As Long
    Dim c As Long
    Dim Compare As Long
    
    For r = 2 To ele
        Compare = num(r)
        For c = r To 1 Step -1
            If c = 1 Then       'If Compare is the lowest number, it will be put at the start of the list
                num(c) = Compare
            ElseIf num(c - 1) > Compare Then    'If Compare is less than the num(c - 1), space is made by moving the greater value to
                num(c) = num(c - 1)             'num(c)
            Else
                num(c) = Compare        'If nothing happens, and Compare is in its correct place, it is placed there, and the loop ends
                c = 1
            End If
        Next c
    Next r
End Sub

Sub Shell_Sort()

'Purpose: The purpose of this subroutine is to sort the data using the Shell Sort method.

'Variable Dictionary

'Variable Name         Scope             Type       Purpose
'k                     Shell_Sort()      Long       This variable is used as a counter in a For Loop.
'i                     Shell_Sort()      Long       This variable is used as a counter in a For Loop.
'c                     Shell_Sort()      Long       This variable is used as a counter in a For Loop.
'r                     Shell_Sort()      Long       This variable is used as a counter in a For Loop.
'Compare               Shell_Sort()      Long       This variable takes the value of the number that will be compared in the
'                                                   Insertion Sort portion of the Shell Sort, so that it can be compared
'                                                   without the number itself being moved, until it needs to be.
'Gap                   Shell_Sort()      Long       This variable contains the value of the current gap being used by the
'                                                   Shell Sort. The programs runs through the Shell Sort eight times, with
'                                                   each gap being smaller than the previous gap.
'Gaps()                Shell_Sort()      Integer    This array contains the values of the gaps that will be used in the
'                                                   Shell Sort.

    Dim k As Long
    Dim i As Long
    Dim c As Long
    Dim r As Long
    Dim Compare As Long
    Dim Gap As Integer
    ReDim Gaps(1 To 8) As Integer
    
    Gaps(1) = 701
    Gaps(2) = 301
    Gaps(3) = 132
    Gaps(4) = 57
    Gaps(5) = 23
    Gaps(6) = 10
    Gaps(7) = 4
    Gaps(8) = 1
     
    For r = 1 To 8
        Gap = Gaps(r)
        For c = 1 To Gap
            For i = c To ele Step Gap       'Same as insertion sort, but with gaps.
                Compare = num(i)
                For k = i To c Step -(Gap)  'The gaps allow numbers to move greater distances in one step towards their
                    If k = c Then           'their proper place, speeding up the sorting process.
                        num(k) = Compare
                    ElseIf num(k - Gap) > Compare Then
                        num(k) = num(k - Gap)
                    Else
                        num(k) = Compare
                        k = c
                    End If
                Next k
            Next i
        Next c
    Next r
              
End Sub
