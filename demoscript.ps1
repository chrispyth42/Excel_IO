#Use excel module and create a new Excel object
#This opens a new instance of Excel in the background to work with on initialization
Using module .\module.psm1
[ExcelIO]$e = [ExcelIO]::new()

#Tell the object to create a new file in the current directory
#All opened workbooks are stored in an array within the object; which the other functions reference. $file is just an int of its index in that array
$file = $e.new([string](Get-Location) + "\demofile.xlsx")

#Populates 3 columns of values to make a table out of, in the opened file $file, on sheet 1
$i = 1
foreach ($cell in $e.Range($file,"A1:A10",1)){
    $cell.Value = $i
    $i = $i + 1
}
$e.Range($file,"B1:B10",1).Value = "=A1*5"
$e.Range($file,"C1:C10",1).Value = "=B1*B1"

#Creates a table out of the populated Cells
#Excel Auto-detects if given a single cell when creating a table
$table = $e.newTable($file,"A1",1)
$table.Name = "DemoTable"

#Names the header columns
$table.HeaderRowRange[1].Value = "Initial Values"
$table.HeaderRowRange[2].Value = "Init *5"
$table.HeaderRowRange[3].Value = "(Init *5)^2"

#Rename the first sheet, and Create a new 2nd sheet
$e.renameSheet($file,1,"Generated Table")
$e.addSheet($file,"Checkerboard")

#Creates a checkerboard within a range by setting background colors
$count = 0
$r = $e.Range($file,1,1,31,31,2) #Ranges can also be specified by (file,x1,y1,x2,y2,sheet)
foreach ($cell in $r){
    if(($count % 2) -eq 0){
        $cell.Interior.ColorIndex = 1
    }
    $count = $count + 1
}

#Sets the width and height of the cells to be roughly square
foreach ($cell in $e.Range($file,1,1,31,1,2)){
    $cell.ColumnWidth = 11/6
}
foreach ($cell in $e.Range($file,1,1,1,31,2)){
    $cell.RowHeight = 10
}

#Save the file and close
$e.save($file)
$e.close()