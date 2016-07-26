Attribute VB_Name = "Heap_Sort"
Option Explicit

'Purpose: The purpose of this module is to contain all the subroutines needed to sort the data using the Heap Sort method.

'Variable Dictionary

'Variable Name         Scope           Type     Purpose
'Heap_Size             General         Long     This variable holds the size of the heap, which is reduced as the data becomes more
'                                               sorted.

Dim Heap_Size As Long


Public Sub Heap_Sort()

'Purpose: The purpose of this subroutine is to sort through the heap, getting it into the order of least to greatest.

'Variable Dictionary

'Variable Name         Scope           Type     Purpose
'Holder                Heap_Sort       Long     This variable holds the value of the number at position 1, so it could be
'                                               switched with another number.
'r                     Heap_Sort       Long     This variable is used as a counter in a For Loop


    Dim Holder As Long
    Dim r As Long
    
    Call Build_Heap
    For r = ele To 2 Step -1    'Puts the heap data in the correct order, from least to greatest.
        Holder = num(1)
        num(1) = num(r)
        num(r) = Holder
        Heap_Size = Heap_Size - 1
        Heapify (1)
    Next r
End Sub

Sub Build_Heap()

'Purpose: The purpose of this subroutine is to build the heap using the heapify function, which will then be sorted in the
'Heap_Sort subroutine.

'Variable Dictionary

'Variable Name         Scope           Type     Purpose
'c                     Build_Heap      Long     This variable is used as a counter in a For Loop

    Dim c As Long

    Heap_Size = ele
    For c = Int(ele / 2) To 1 Step -1   'Builds the heap by heapifying the first half of the data, which creates a full heap.
        Heapify (c)
    Next
End Sub


Function Heapify(i As Long)

'Purpose: The purpose of this function is to put the three values, the parent value, and its two children, in the correct
'order for the Heap data structure.

'Variable Dictionary

'Variable Name         Scope           Type     Purpose
'Largest_Index         Heapify         Long     This variable holds the position, or index, of the largest value of the
'                                               three compared in the Heapify function.
'Largest_Value         Heapify         Long     This variable holds the value of the largest value of the three compared
'                                               in the Heapify function.
'Parent_Index          Heapify         Long     This variable holds the position, or index, of the parent value, which the
'                                               function then uses to find the two child values and indexes.
'Parent_Value          Heapify         Long     This variable holds the value of the parent value.
'LChild_Index          Heapify         Long     This variable holds the position, or index, of the left child value, which is
'                                               twice the parent index, if it exists.
'LChild_Value          Heapify         Long     This variable holds the value of the left child of the parent value, if it exists.
'LChild_Exist          Heapify         Boolean  This boolean is used so that, if the left child doesn't exist, the program will
'                                               not crash trying to find it.
'RChild_Index          Heapify         Long     This variable holds the position, or index, of the right child value, which is
'                                               twice the parent index plus one, if it exists.
'RChild_Value          Heapify         Long     This variable holds the value of the right child of the parent value, if it exists.
'RChild_Exist          Heapify         Boolean  This boolean is used so that, if the right child doesn't exist, the program will
'                                               not crash trying to find it.

    Dim Largest_Index As Long
    Dim Largest_Value As Long
    Dim Parent_Index As Long
    Dim Parent_Value As Long
    Dim LChild_Index As Long
    Dim LChild_Value As Long
    Dim LChild_Exist As Boolean
    Dim RChild_Index As Long
    Dim RChild_Value As Long
    Dim RChild_Exist As Boolean
    LChild_Exist = False
    RChild_Exist = False
    Parent_Index = i
    Largest_Index = i
    Largest_Value = num(i)
    If (2 * i) <= Heap_Size Then    'Checks if LChild_Index exists
        LChild_Index = 2 * i
        LChild_Exist = True
    End If
    If (2 * i + 1) <= Heap_Size Then    'Checks if RChild_Index exists
        RChild_Index = 2 * i + 1
        RChild_Exist = True
    End If
    Parent_Value = num(i)
    If LChild_Exist = True Then
        LChild_Value = num(2 * i)
        If Parent_Value < LChild_Value Then     'Checks if Parent and LChild are in correct order
            Largest_Index = LChild_Index        'If so, then change the largest value and index to LChild's
            Largest_Value = num(2 * i)
        End If
    End If
    If RChild_Exist = True Then
        RChild_Value = num((2 * i) + 1)
        If Largest_Value < RChild_Value Then    'Checks if RChild is greater than the largest value
            Largest_Index = RChild_Index        'If so, then change the largest value and index to RChild's
            Largest_Value = RChild_Value
        End If
    End If
    If Largest_Index <> i Then      'Checks if anything should be switched
        num(i) = Largest_Value      'Switch largest value and parent value
        num(Largest_Index) = Parent_Value
        Heapify (Largest_Index)     'If anything is switched, the heapify function runs again
    End If
End Function

