# we require the openxlsx package
if (!require(openxlsx)) install.packages("openxlsx")
library(openxlsx)

# The set style function is to define stuffs (colors, and
# so on for your sheet)
#
# Feel free to change if you want

set_style <- function() {
  options("openxlsx.borderstyle" = "thick")
  options("openxlsx.borderColour" = "#4F80BD") # The border colors


  # Style of the header of the tables. ?createStyle for openxlsx
  hs <- createStyle(
    fontColour = "black",
    fgFill = "#DCE6F1",
    halign = "center",
    valign = "center",
    textDecoration = "Bold",
    border = "TopBottomLeftRight",
    borderStyle = "thin"
  )


  # The style of a title of the sheet
  ts <- createStyle(
    fontColour = "#1e49f9",
    fgFill = "#DCE6F1",
    halign = "left",
    valign = "center",
    textDecoration = "Bold",
    border = "TopBottom",
    fontSize = 14,
    borderStyle = "medium"
  )

  # The style of the name of the table
  ns <- createStyle(
    fontColour = "#1e49f9",
    halign = "left",
    valign = "center",
    fontSize = 12,
    textDecoration = "bold"
  )

  sheetsh <- createStyle(
    halign = "left",
    valign = "center",
    wrapText = TRUE
  )

  return(list(hs = hs, ts = ts, ns = ns, sheetsh = sheetsh))
}

# The purpose of the write at head is to put the header names for each
# sheet of a workbook.
#
write_at_head <- function(wb, sheetnames, headers) {
  st <- set_style()

  # I really didn't know how to use the
  #* apply functions with that.

  for (i in 1:length(sheetnames)) {
    writeData(wb,
      sheet = sheetnames[i],
      x = headers[i],
      startCol = 2,
      startRow = 2,
      colNames = FALSE,
      rowNames = FALSE,
      borders = "rows",
      borderStyle = "thick"
    )
  }

  lapply(sheetnames, function(x) {
    addStyle(wb,
      x,
      style = st$ts,
      rows = 2,
      cols = 2:5
    )
    setColWidths(
      wb,
      x,
      cols = 1:200,
      widths = 35
    )
    setRowHeights(
      wb,
      x,
      rows = 1:10000,
      heights = 25
    )
    addStyle(wb,
      x,
      style = st$sheetsh,
      rows = 1:10000,
      cols = 1:200,
      gridExpand = TRUE,
      stack = TRUE
    )
  })
}

# The purpose of the function is to just copy the
# create workbook function with a huge title at
# each begining of each sheet.

initiate_workbook <- function(sheetnames, headernames) {
  # more lisible
  if (length(sheetnames) != length(headernames)) stop("Sheetnames and headernames must be same length")

  if (!(is.character(sheetnames) & is.character(headernames))) stop("Sheetnames and headernames must be character vectors")

  wb <- createWorkbook()

  # Change the overall font if you want
  modifyBaseFont(wb,
    fontSize = 9,
    fontColour = "black",
    fontName = "Calibri"
  )

  lapply(
    sheetnames,
    function(x) {
      addWorksheet(wb,
        x,
        gridLines = FALSE,
        zoom = 95
      )
    }
  )

  write_at_head(wb, sheetnames, headernames)
  return(wb)
}

# writing one table to one spread sheet
# I need to know at which line I have to start writing,
# and where I stopped.

push_one_table <- function(data, wb, sheetname, tablename,
                           tablelabel,
                           previous_stop = 4) {
  st <- set_style()

  writeData(wb,
    sheet = sheetname, x = tablelabel,
    startRow = previous_stop,
    startCol = 2
  )

  # Here I add the header style to the name of the table
  addStyle(wb,
    sheet = sheetname,
    style = st$ns,
    rows = previous_stop,
    cols = 2
  )

  previous_stop <- previous_stop + 2

  # writing the data itself
  writeDataTable(wb,
    sheet = sheetname,
    x = data,
    startCol = 2,
    startRow = previous_stop,
    colNames = TRUE,
    rowNames = FALSE,
    tableName = tablename,
    tableStyle = "TableStyleMedium2"
  )

  previous_stop <- previous_stop + nrow(data) + 3

  return(previous_stop)
}

# The end is to push all tables of table list and to add the names
# in the excel sheet.

push_all_tables <- function(wb,
                            listoftables,
                            listofnames,
                            listoflabels,
                            sheetname) {
  if (length(listoftables) != length(listofnames)) stop("Sheetnames and headernames must be same length")


  stop <- 4

  for (i in 1:length(listoftables)) {
    new_stop <- push_one_table(
      as.data.frame(listoftables[[i]]),
      wb,
      sheetname = sheetname,
      tablename = listofnames[i],
      tablelabel = listoflabels[i],
      previous_stop = stop
    )

    stop <- new_stop
  }
}
