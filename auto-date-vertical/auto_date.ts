function main(workbook: ExcelScript.Workbook) {
    
    // get the new date and format the date to mm//dd//yyy
    const currentDate = new Date();
    const formattedDate = currentDate.toLocaleDateString('en-US');

    //gets the currently selected range
    const selectedRange: ExcelScript.Range = workbook.getSelectedRange();

    // check to ensure a single cell was selected
    if (selectedRange.getRowCount() !== 1 || selectedRange.getColumnCount() !==1) {
        console.log("Please Select a single cell...")
    }

    const {startRow, startColumn} = getStartIndices(selectedRange);


    //this is the loop that fills the date into a custom number of boxes 
    // change the "insert number here" to how many time you need the date filled
    for (let i = 0; i < "insert number here"; i++ ) {
        selectedRange.getWorksheet().getRangeByIndexes(startRow + i, startColumn, 1, 1).setValue(formattedDate);
    }

}

//helper method to get the starting row and column indices
function getStartIndices(range: ExceScript.Range): {startRow: number, startColumn: number } {
    const startRow = range.getRowIndex();
    const startColumn = range.getColumnIndex();
    return { startRow, startColumn };
}