# Excel-Data-Preparation-AddIn
We have two option
1. Keep Rows that validate a condition
2. Remove Rows that validate a condition.

And below are the list of conditions:

|Name |Description|
|---|---|
|Contains Search Text|Contains User Specified Info.|
|Does Not Contains Search Text|Does Not Contains User Specified Info.|
|Start With Search Text|Start With User Specified Info .|
|Does Not Start With  Search Text|Does Not Start With User Specified Info .|
|End With Search Text|Ends With User Specified Info .|
|Does Not End With Search Text|Does Not Ends With User Specified Info .|
|Search Characters In Anywhere|Has All Characters Of Search Text In Anywhere.|
|Search Characters In Sequence|Has All Character Of Search Text In Sequence.|
|All Character In Uppercase|All Character In Uppercase|
|All Character In Lowercase|All Character In Lowercase|
|All Character In Proper Case|All Character In Proper Case|
|Has More Than Chars Count|Has More Than Specified Number Of  Character.|
|Has Number Of Chars| Has Specific Number Of Characters.|
|At Least One Delimiter Present|Has At Least One Delimiter.|
|All Delimiter Present|Has All Delimiter.|
|Between Two Value|Are In Between Two Value.|
|Does Not Between Two Value|Are Not  Between Two Value.|
|Greater Than Specified Value|Greater Than Given Value.|
|Less Than Specified Value|Less Than Given Value.|
|Divisible By Specified Number|Are Divisible By Number.|
|Even Numbers|Even Numbers.|
|Odd Numbers|Odd Numbers.|
|Prime Number|Is Prime Number.|
|Only Text|Only Contains Text.|
|Only Numbers|Only Contains  Numbers.|
|Only Alphanumeric|Only Has Alphanumeric[a-z,0-9] One.|

# AddIn Installation:
See this to know about [How To Install A AddIn](https://www.youtube.com/watch?v=qc-HjrAscQ8)

# Warning

This AddIn uses CurrentRegion when running any specific action. It also clears out content from CurrentRegion and updates with result data. So make sure you keep a backup of your data or run at your own risk.

# Use Cases

We have a [Test Data](Test%20Data.xlsm) File.  

Search Characters In Anywhere:{"Search Text":"2w9","ColIndex":3} >> This means use Search Characters In Anywhere button and use 2w9 as Search Text and Keep your cursor in the third column before running this command.  

[Here](https://youtu.be/cG3d_TdLGN4) is a video demo of this AddIn.