Read this B4 you use.

While using MSHFlex Grid you must set colwidth to each column. 
See load event of the form, colsetupcode method also generates the 
code needed but you must slide every column.

If you issued "Flex.ColWidth(3) = 0 " column#3 does not print.
If you issued "Flex.RowHeight(3) = 0 " row#3 does not print.

If you try to make rowheight smaller than text, it looks ok in 
picture box but while printing it prints over the cell one below that.
So give enough rowsheight.

Value of RowFrom should always be smaller than RowTo.
CurX and CurY are like CurrentX ans CurrentY of Vb.

Printing is quite nice in my laserjet 6L. 

I am just an ordinary one, these codes may still
contain errors. If you find one inform me I will try to fix it.

Do send your comments, suggestions and improvement.
This much for now.

Opal Raj Ghimire
Kathmandu, Nepal
Updated 22nd Dec 2001