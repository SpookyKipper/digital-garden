---
{"dg-publish":true,"permalink":"/king-ling/ict-test/","created":"2024-06-17T14:04:45.889+08:00","updated":"2024-06-17T19:00:25.666+08:00"}
---

#exam #ict
## Excel
### =sum(range)
**range (array)**
### =sum(number1, \[number2], ...)
**numberX (cell ref)**
### =count(range)
**range (array)**
### =count(number1, \[number2], ...)
**numberX (cell ref)**
### =sumif(range, criteria, \[sum_range])
**range (cell ref)**
	where it matches criteria
**criteria (string)**
	what to expect in range
**sum_range (cell ref array)**
	relative to range, actual cells to use
### =sumif(range, criteria)
**range (cell ref)**
	where it matches criteria
**criteria (string)**
	what to expect in range
### =rank(number, ref, \[order])
**number (integer)**
**ref (array)**
**order (boolean)**
	0: Descending
	1: Ascending

### =xlookup(lookup_value, lookup_array, return_array, \[if_not_found], \[match_mode], \[search_mode])
**lookup_value (string/number)**
**lookup_array (array)**
	where lookup_value is in
**return_array (array)**
	where the value corresponds to lookup_value
**if_not_found (string/number)**
	what to echo when lookup_value does not exist
**match_mode (integer)**
	0: Exact match <u>DEFAULT</u>
	~~-1: Exact match or next smaller item (sorting required)~~
	~~1: Exact match or next larger item (sorting required)~~
	~~2: Wildcard character match (\*\?\~)~~
**search mode (integer)**
	~~1: Perform a search starting at the first item. <u><s>DEFAULT</s></u>
	-1: Perform a reverse search starting at the last item.
	2: Perform a binary search that relies on lookup_array being sorted in ascending order. If not sorted, invalid results will be returned.
	-2: Perform a binary search that relies on lookup_array being sorted in descending order. If not sorted, invalid results will be returned.~~

