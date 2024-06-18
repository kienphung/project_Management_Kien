class ExcelManager {
  constructor() {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        this.initialize();
      }
    });
  }

  async initialize() {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.onSelectionChanged.add(this.handleSelectionChange.bind(this));
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }

  async handleSelectionChange(event) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = context.workbook.getSelectedRange();
        range.load("address, rowIndex, columnIndex, values");
        await context.sync();

        // Check if the selected cell is in column A
        if (range.columnIndex === 0) {
          const rowIndex = range.rowIndex + 1; // Excel row index is 1-based
          const rowRange = sheet.getRange(`A${rowIndex}:Z${rowIndex}`);
          rowRange.load("values");
          await context.sync();

          const rowData = rowRange.values[0];
          this.populateForm(rowData, rowIndex);
        }
      });
    } catch (error) {
      console.error(error);
    }
  }

  populateForm(rowData, rowIndex) {
    // Assuming the form fields have IDs rowIndex and rowData
    document.getElementById("rowIndex").value = rowIndex;
    document.getElementById("rowData").value = rowData.join(", ");
  }

  generateUniqueId() {
    return `ID-${Date.now()}`;
  }

  async addRow(rowIndex) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const newRowRange = sheet.getRange(`A${rowIndex}`).getEntireRow();
        newRowRange.insert(Excel.InsertShiftDirection.down);

        const uniqueId = this.generateUniqueId();
        const idCell = sheet.getRange(`B${rowIndex}`);
        idCell.values = [[uniqueId]];

        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }

  async updateRow(rowIndex, rowData) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(`A${rowIndex}:Z${rowIndex}`);
        range.values = [rowData];
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }

  async deleteRow(rowIndex) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(`A${rowIndex}:Z${rowIndex}`);
        range.delete(Excel.DeleteShiftDirection.up);
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }
}

// Helper functions to get values from form
function getRowIndex() {
  return parseInt(document.getElementById("rowIndex").value);
}

function getRowData() {
  return document.getElementById("rowData").value.split(", ");
}

// Instantiate ExcelManager
const excelManager = new ExcelManager();

// Bind form buttons to class methods
document.getElementById('viewButton').addEventListener('click', () => excelManager.viewRow(getRowIndex()));
document.getElementById('addButton').addEventListener('click', () => excelManager.addRow(getRowIndex()));
document.getElementById('updateButton').addEventListener('click', () => excelManager.updateRow(getRowIndex(), getRowData()));
document.getElementById('deleteButton').addEventListener('click', () => excelManager.deleteRow(getRowIndex()));
