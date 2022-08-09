const WHITE = "ffffff";
const BLACK = "000000";
const GREY = "bfbfbf";

function main(workbook: ExcelScript.Workbook)
{
  // get the selected cell and check whether it's our white month label
  // not foolproof, but checking for the date format should be good enough
  let cell = workbook.getSelectedRange();
  if (cell.getNumberFormat() != 'mmm-yy')
  {
    console.log("Please select the large white date cell at the top of the month before the month you wish to generate!");
    return;
  }

  // select all cells above and below it
  let range = cell.getEntireColumn();

  // create a new set of cells and duplicate the month to it
  // this actually creates new cells to the LEFT of the selected month,
  // pushing our selection to right which we will edit into the new month
  range.insert(ExcelScript.InsertShiftDirection.right).copyFrom(range);

  // get the cell address - we'll reference it when creating dates later
  let address = cell.getAddress();

  // change the date at the top of the month
  let value = Number(cell.getValue());
  let date = new Date(1900, 0, value - 1); // get the date of the selected month

  date.setMonth(date.getMonth() + 1); // add a month
  date.setDate(date.getDate() - 1); // go back one day so the date is the last day of the selected month
  cell.setValue(value + date.getDate()); // add the amount of time in the month to the cell's value - this will change it to the start of the next month
  date.setDate(date.getDate() + 1); // go forward one day so this date represents the start of the next month again

  // clear all contents and fill the month with the green background color
  let format = cell.getOffsetRange(1, 0).getFormat();
  let color = format.getFill().getColor();

  cell = cell.getOffsetRange(2, 0);
  range = cell.getResizedRange(90, 0);
  range.clear();

  format = range.getFormat();
  format.getFill().setColor(color);
  format.getFont().setSize(8);
  format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  // the exciting part!! we need to generate the days in each month
  // the first step is to figure out which row to start on
  // if the first day of the month is a Monday, we start on the first row,
  // but we will need to skip rows if it is any other day of the week
  let weekday = 1; // 1 is Monday
  let startDay = date.getDay();

  while (weekday != startDay)
  {
    cell = skipRows(cell, weekday);
    weekday = (weekday + 1) % 7;
  }
  range = cell;

  // fill 'em in!
  let month = date.getMonth();
  let columns = range.getColumnCount();
  while (month == date.getMonth())
  {
    let day = date.getDate();
    weekday = date.getDay();

    range = range.getAbsoluteResizedRange(getWeekdayHeight(weekday), columns);
    format = range.getFormat();

    if (isWeekend(weekday)) // weekends are bold and grey
    {
      format.getFill().setColor(GREY);
      format.getFont().setBold(true);
    }
    else // weekdays are white
    {
      format.getFill().setColor(WHITE);
    }

    format.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor(BLACK);
    format.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor(BLACK);
    format.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor(BLACK);
    format.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor(BLACK);

    cell = range.getCell(0, 0);
    cell.setValue("=" + address + "+" + (day - 1));
    cell.setNumberFormat("d");

    date.setDate(day + 1);
    range = skipRows(range, weekday);
  }
}

function skipRows(range: ExcelScript.Range, weekday: number)
{
  return range.getOffsetRange(getWeekdayHeight(weekday), 0);
}

function getWeekdayHeight(weekday: number)
{
  if (isWeekend(weekday))
    return 1;
  
  return 3;
}

function isWeekend(weekday: number)
{
  return weekday == 0 || weekday == 6; // 0 is Sunday, 6 is Saturday
}
