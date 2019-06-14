class ExcelIO {
    ########################################################################################
    #Main Variables and constructor
    ########################################################################################
    #Excel Application itself
    [System.__ComObject]$app
    #Array of all opened excel files
    [array]$books

    #Constructor: opening Excel in invisible mode with all alerts suppressed
    ExcelIO(){
        $this.app = New-Object -ComObject Excel.Application
        $this.app.Visible = $false
        $this.app.DisplayAlerts = $false
    }

    ########################################################################################
    #Application Management
    ########################################################################################
    #Quit excel application
    [void]close(){
        #Closes the workbooks, then Application
        $this.app.Workbooks.Close()
        $this.app.Quit()

        #Releases the COM object for the excel application, and then for each of the workbooks
        #So that the Excel process properly exits on close
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.app)
        foreach($file in $this.books){
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($file)
            $file = $null
        }
        $this.app = $null

        #Run garbage collection
        [GC]::Collect()
    }

    ########################################################################################
    #File Management
    ########################################################################################
    #Saves an opened workbook in the books array
    [void]save([int]$wb){
        try{
            $this.books[$wb].Save()
        }catch{
            Write-Host(".save: failed to save workbook $wb")
        }
    }

    #Saves all open files
    [void]saveAll(){
        foreach($file in $this.books){
            $file.Save()
        }
    }

    #Opens an excel file, adding a reference to it in the Array 'books'
    [int]open([string]$path){
        try{
            $wb = $this.app.Workbooks.Open($path)
            $this.books += $wb
            return ($this.books.Length - 1)
        }catch{
            Write-Host(".open: File does not exist, or is invalid: $path")
            return -1
        }
    }
    #Same as .open, but read only 
    [int]openReadOnly([string]$path){
        try{
            $wb = $this.app.Workbooks.Open($path,$null,$true)
            $this.books += $wb
            return ($this.books.Length - 1)
        }catch{
            Write-Host(".openReadOnly: File does not exist, or is invalid: $path")
            return -1
        }
    }

    #Uses the builtin windows dialog to open excel files
    #I didn't write the first half of this, and got it from the page below
    #https://code.adonline.id.au/folder-file-browser-dialogues-powershell/
    [array]openDialog(){
        #Uses the builtin windows open-file dialog
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Multiselect = $true # Multiple files can be chosen
            Filter = 'Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx' # Specified file types
        }
        
        #Getting filenames
        [void]$FileBrowser.ShowDialog()
        $path = $FileBrowser.FileNames;
        
        #If the user selected something, iterate through and open the selected files; returning an array of indexes to those opened files
        if ("" -ne $path){
            $opened = @()
            foreach ($file in $path){
                $opened += $this.open($file)
            }
            return $opened
        #If the user selected nothing, write that, and return empty array
        }else{
            Write-Host (".openDialog: Open file cancelled by user")
            return @()
        }
    }

    #Creates a new file, and saves it to specified file location; adding it to array of books
    [int]new([string]$path){
        #Check if Path exists to avoid overwrite
        if(!(Test-Path $path)){    
            try{
                $newFile = $this.app.Workbooks.Add()
                $newFile.SaveAs($path)
                $this.books += $newFile
                return ($this.books.Length - 1)
            }catch{
                Write-Host(".new: Failed to save new file (Probably invalid path, or inadequite permissions): $path")
                return -1
            }
        }else{
            Write-Host(".new: File already exists at location: $path")
            return -1
        }
    }

    ########################################################################################
    #Sheets
    ########################################################################################
    #Adds a sheet to the workbook, to the right of other sheets
    #https://social.technet.microsoft.com/Forums/azure/en-US/d89913a1-c1bf-49b0-b76f-703740fb5aba/add-or-move-excel-sheet-with-powershell?forum=ITCG
    [void]addSheet([int]$wb,[string]$name){
        try{
            #I don't understand why this line works, but I got this from the link Above: Creates a new sheet to the right of all other sheets in a workbook
            $newSheet = $this.books[$wb].Worksheets.Add([System.Reflection.Missing]::Value,$this.app.Worksheets.Item($this.app.Worksheets.count))
            if($name -ne ""){
                $newSheet.Name = $name
            }
        }catch{
            Write-Host(".addSheet: Workbook $wb doesn't exist")
        }
        
    }
    [void]addSheet([int]$wb){ $this.addSheet($wb,"") }

    #Gets a count of how many sheets are in a workbook
    [int]countSheets([int]$wb){
        try{ 
            $output = 0
            foreach ($sheet in $this.books[$wb].Worksheets){ $output = $sheet.Index }
            return $output
        }catch{
            return -1
            Write-Host(".countSheets: Workbook $wb doesn't exist")
        }
    }

    #Accepts the index of a sheet in a workbook, and renames it with a string
    [void]renameSheet([int]$wb,[int]$sheet,[string]$name){
        try{
            $s = $this.books[$wb].Worksheets.Item($sheet)
            $s.name = $name
        }catch{
            Write-Host(".renameSheet: Sheet or Workbook doesn't exist (wb $wb, sheet $sheet)")
        }
    }

    ########################################################################################
    #Cell/Range Accessors
    ########################################################################################
    #Returns a link to a Cell object, or Null if it fails
    #Is overloaded to accept either a coordinate pair, or a Range like "A1"
    [System.Object]cell([int]$wb,[int]$x,[int]$y,[int]$sheet){
        try{
            return $this.books[$wb].Sheets.Item($sheet).Cells($y,$x)
        }
        catch{
            Write-Host(".cell: Failed to get cell (Workbook $wb, x $x, y $y, sheet $sheet)")
            return $null
        }
    }
    [System.Object]cell([int]$wb,[string]$r,[int]$sheet){
        try{
            return $this.range($wb,$r,$sheet)[1]
        }
        catch{
            Write-Host(".cell: Failed to get cell (Workbook $wb, range $r, sheet $sheet)")
            return $null
        }
    }

    #Returns a link to a Range of cells, or Null if it fails
    #Accepts either a coordinate pair, or a value like "A1:B10"
    [System.Object]range([int]$wb,[int]$x1,[int]$y1,[int]$x2,[int]$y2,[int]$sheet){
        $r = $this.convertToA1($x1,$y1,$x2,$y2)
        return $this.range($wb,$r,$sheet)
    }
    [System.Object]range([int]$wb,[string]$range,[int]$sheet){
        try{
            #Retreives the range from the workbook
            return $this.books[$wb].Sheets.Item($sheet).Range($range)
        }catch{
            Write-Host(".range: Failed to get range (Workbook $wb, range $range, sheet $sheet)")
            return $null
        }
    }

    ########################################################################################
    #Tables
    ########################################################################################
    #Creates a new table at a specified location, and returns the table object to the user
    #Either by providing coordinate pair/pairs, or a Range
    [System.Object]newTable([int]$wb,[int]$x1,[int]$y1,[int]$x2,[int]$y2,[int]$sheet){
        $r = $this.convertToA1($x1,$y1,$x2,$y2)
        return $this.newTable($wb,$r,$sheet)
    }
    [System.Object]newTable([int]$wb,[int]$x,[int]$y,[int]$sheet){
        $r = $this.convertToA1($x,$y)
        return $this.newTable($wb,$r,$sheet)
    }
    [System.Object]newTable([int]$wb,[string]$r,[int]$sheet){
        try{
            #Gets the sheet from the workbook, select the desired range, and Add a new table in it
            $s = $this.books[$wb].Sheets.Item($sheet)
            $s.Range($r).Select()
            return $s.ListObjects.Add()
        }catch{
            Write-Host(".newTable: Failed to create table at $r (Workbook or Sheet doesn't exist, or table is overlapping another table)")
            return $null
        }
    }

    #Retrieves a table by name from a workbook/sheet
    [System.Object]getTable([int]$wb,[string]$table,[int]$sheet){
        try{
            return $this.books[$wb].Sheets.Item($sheet).ListObjects._Default($table)
        }catch{
            Write-Host(".getTable: Failed to retrieve Table (Workbook $wb, Table '$table', Sheet $sheet")
            return $null
        }
    }

    #Adds a row to that table, and returns the Range of the added row so that the scripter can do things with it
    #Can be used by targetting a table within a workbook/sheet, or by handing it a table directly
    [System.Object]addTableRow([int]$wb,[string]$table,[int]$sheet){
        try{
            $this.books[$wb].Sheets.Item($sheet).ListObjects._Default($table).ListRows.Add(1)
            return  $this.books[$wb].Sheets.Item($sheet).ListObjects._Default($table).ListRows(1).Range
        }catch{
            Write-Host(".addRow: Failed to add row to Table (Workbook $wb, Table '$table', Sheet $sheet")
            return $null
        }
    }
    [System.Object]addTableRow([System.Object]$table){
        try{
            $table.ListRows.Add(1)
            return $table.ListRows(1).Range
        }catch{
            Write-Host(".addRow: Failed to add row to given table")
            return $null
        }   
    }

    ########################################################################################
    #Utilities
    ########################################################################################

    #Converts Integers into A1 excel column letter format
    #I didn't write this, and got it from this page here
    #https://gallery.technet.microsoft.com/office/Powershell-function-that-88f9f690
    [string]convertToA1([int]$number){ 
        $a1Value = $null 
        While ($number -gt 0) { 
          $multiplier = [int][system.math]::Floor(($number / 26)) 
          $charNumber = $number - ($multiplier * 26) 
          If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 } 
          $a1Value = [char]($charNumber + 64) + $a1Value 
          $number = $multiplier 
        } 
        Return $a1Value 
      }
    #Overloads that I wrote to fully convert coordinates/coordinate pairs to Excel A1
    [string]convertToA1([int]$x,[int]$y){
        return ( $this.convertToA1($x) + $y )
    }
    [string]convertToA1([int]$x1,[int]$y1,[int]$x2,[int]$y2){
        if(($x1 -eq $x2) -and ($y1 -eq $y2)){
            return $this.convertToA1($x1,$y1)
        }else{
            return $this.convertToA1($x1) + $y1  + ":" + $this.convertToA1($x2) + $y2 
        }
    }
}