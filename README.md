# CodeSnippits
This is a collection of simple standalone code for a specific purpose.

## (File:ConsecutiveRef.bas) Adding En Dash Automatically for Consecutive Numbers
This is just a word macro. Apparently, some writers who use references may need this. After selecting the text and running the Macro, the Macro find strings of comma seperated numbers and inserts an en dash between only the first and the last of consecutive numbers in that string removing the full series.
Basically, it should convert this:  
Berlin 11, 12, 13, 14, 15, 16, 18, 23, 24, 25, 29  
Paris 50, 51, 52, 53, 54, 55, 60, 61
  
Into this:  
Berlin 11–16, 18, 23–25, 29  
Paris 50–55, 60–61
