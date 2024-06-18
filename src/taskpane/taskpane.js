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
        const sheet = context.workbook.worksheets.getItem("GanttChart");
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
        const sheet = context.workbook.worksheets.getItem("GanttChart");
        const range = context.workbook.getSelectedRange();
        range.load("address, rowIndex, columnIndex, values");
        await context.sync();

        // Check if the selected cell is in column A
        if (range.columnIndex === 0 && range.rowIndex >= 7) {
          const rowIndex = range.rowIndex + 1; // Excel row index is 1-based
          const rowRange = sheet.getRange(`A${rowIndex}:J${rowIndex}`);
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
    document.getElementById("rowIndex").value = rowIndex;
    document.getElementById("rowData").textContent = rowData[0];
    const taskName = document.getElementById("rowTask");
          taskName.innerText = rowData[1];
    document.getElementById("rowCTiet").value = rowData[9];
    this.showLinks(rowData,rowIndex,null);
  }
  textToArray(text) {
    console.log(text);
    // Loại bỏ dấu ngoặc vuông và khoảng trắng không mong muốn
      //kiem tra neu chuoi rong
    if(typeof text === 'string'){
      if(text ===''){
        return [''];
      }else if(!text.match(/,/)){
        //console.log("khong bao gom dau ','");
        return [text];
      }else{

        const regex = /[\[\]\s]/g;
        if(regex.test(text)){
          const cleanedText = text.replace(regex, '');
          // Chia chuỗi thành mảng các phần tử
          const array = cleanedText.split(',');
          return array;
        } else {
          //console.log("bao gom dau ','");
          // Chia chuỗi thành mảng các phần tử
          const array = text.split(',');
          return array;
        }
      }

    }else{
      return text;
    }
    
}
  async showLinks(rowData,rowIndex,data){
    try{
      //console.log(rowData);
      //console.log(data);
      const fileLink = document.getElementById("fileLink");
      //const a = rowData[8];
      let rData = [];
      // Xóa tất cả các phần tử con hiện có
      fileLink.innerHTML = '';
      if(rowData && data=== null){
        rData = this.textToArray(rowData[8]);
      }else if(data && rowData === null){
        rData = this.textToArray(data);
        //console.log(rData);
      }
    //console.log(rData);

    if(Array.isArray(rData)){

        rData.forEach((el, index) => {
          let count = index;
          
          // Tạo một phần tử con, ví dụ <p>
          const p = document.createElement('p');
          // Đặt nội dung của phần tử con
          const regex = /\((.*?)\)\((.*?)\)/g;
          if (regex.test(el)) {
            const replacedText = el.replace(regex, '<a href="$2" target="_blank">$1</a>');
          p.innerHTML= '<span class="badge bg-danger" id="'+count+'"> </span>'+ replacedText
          } else {
            p.innerHTML='<span class="badge bg-danger" id="'+count+'"> </span> <a href="'+el+'" target="_blank">'+ el+'</a>'
          }        
          // Thêm sự kiện click vào mỗi phần tử
          p.addEventListener('click', (event) => {
            const clickedId = event.target.id; // Lấy id của phần tử vừa click
            if(clickedId){
              this.removeLink(rData,index,rowIndex);
            }
          });
          // Thêm phần tử con vào phần tử <div>
          fileLink.appendChild(p);
        });
    }

  
    //document.getElementById("fileLink").innerHTML = (replacedText !==rData)?'<p>'+replacedText+'</p>':'<p><a href="'+rData+'">'+rData+'</a></p>';
    } catch(error){
      console.log(error);
    }
  }
  //Remove Link Files
  
  async removeLink(rowData,id,rowIndex){
    
    //remove el ra khỏi mảng của cell   
    rowData.splice(id, 1);
    const text = rowData.join(", ");
    //update vào cell
    await Excel.run(async (context) => {     
      const sheet = context.workbook.worksheets.getItem("GanttChart");

    // Lấy ô cần cập nhật
    const range = sheet.getRangeByIndexes(rowIndex-1, 8, 1, 1);
    range.load("values");
    await context.sync();

    // Lấy dòng chứa ô
    const row = range.getEntireRow();
    row.load("format/rowHeight");
    await context.sync();
    // Lấy chiều cao mặc định của dòng
    const defaultRowHeight = row.format.rowHeight;  
  
    // Gán dữ liệu mới vào ô
    range.values = [[text]];
   
      // Kiểm tra và điều chỉnh chiều cao của dòng nếu cần
      //if (row.format.rowHeight > defaultRowHeight) {
        row.format.rowHeight = defaultRowHeight;
      //}

  await context.sync();

  });
    this.showLinks(null,rowIndex,rowData);
  }
  // End remove link files
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
  };


  async updateRow(rowIndex, rowCTiet) {
    //try {
      await Excel.run(async (context) => {     
          const sheet = context.workbook.worksheets.getItem("GanttChart");

        // Lấy ô cần cập nhật
        const range = sheet.getRangeByIndexes(rowIndex-1, 9, 1, 1);
        range.load("values");
        await context.sync();

        // Lấy dòng chứa ô
        const row = range.getEntireRow();
        row.load("format/rowHeight");
        await context.sync();
        // Lấy chiều cao mặc định của dòng
        const defaultRowHeight = row.format.rowHeight;  
      
        // Gán dữ liệu mới vào ô
        range.values = [[rowCTiet]];
       
          // Kiểm tra và điều chỉnh chiều cao của dòng nếu cần
          //if (row.format.rowHeight > defaultRowHeight) {
            row.format.rowHeight = defaultRowHeight;
          //}

      await context.sync();



      });
    //} catch (error) {
   //   console.error(error);
   // }
  }
  async addLinkRow(rowIndex, rowLink) {
    //try {
      await Excel.run(async (context) => {
      
      
          const sheet = context.workbook.worksheets.getItem("GanttChart");

        // Lấy ô cần cập nhật
        const range = sheet.getRangeByIndexes(rowIndex-1, 8, 1, 1);
        range.load("values");

        await context.sync();

        // Lấy dòng chứa ô
        const row = range.getEntireRow();
        row.load("format/rowHeight");
        await context.sync();
        // Lấy chiều cao mặc định của dòng
        const defaultRowHeight = row.format.rowHeight;  
        if(range.values[0][0]==''){
          // Gán dữ liệu mới vào ô
          range.values = [[rowLink]];
          this.showLinks(null,rowIndex,range.values[0][0]);
          document.getElementById("rowLink").value='';
         // this.addLinkToHtml(rowLink);
        }else{
          //console.log("cộng thêm array");
          
          const b = this.textToArray(range.values[0][0]);
          const a = [...b,rowLink];
          const c = a.join(',');
          //console.log(rowLink);
          range.values =[[c]];

          this.showLinks(null,rowIndex,range.values[0][0]);
          //this.addLinkToHtml(range.values[0][0]);
          document.getElementById("rowLink").value='';
        }
        
        
       
          // Kiểm tra và điều chỉnh chiều cao của dòng nếu cần
          //if (row.format.rowHeight > defaultRowHeight) {
            row.format.rowHeight = defaultRowHeight;
          //}

      await context.sync();



      });
    //} catch (error) {
   //   console.error(error);
   // }
  }
  async addLinkToHtml(link){
   // const tempLink = convertToLinks(link);
   try{
    document.getElementById("rowLink").value="";
        const regex = /\[(.*?)\]\((.*?)\)/g;
        const replacedText = link.replace(regex, '<a href="$2">$1</a>');
        const uplink = document.getElementById("fileLink");
              uplink.innerHTML = (replacedText!==link)?'<p>'+ replacedText + '</p>':'<p><a href="'+link+'">'+link+'</a></p>';
   } catch (error) {
      console.log(error);
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
function getRowCTiet(){
  return document.getElementById("rowCTiet").value;
}
function getRowLink(){
  return document.getElementById("rowLink").value;
}
// Instantiate ExcelManager
const excelManager = new ExcelManager();

// Bind form buttons to class methods
//document.getElementById('viewButton').addEventListener('click', () => excelManager.viewRow(getRowIndex()));
//document.getElementById('addButton').addEventListener('click', () => excelManager.addRow(getRowIndex()));
document.getElementById('updateButton').addEventListener('click', () => excelManager.updateRow(getRowIndex(), getRowCTiet()));
document.getElementById('linkButton').addEventListener('click', () => excelManager.addLinkRow(getRowIndex(), getRowLink()));
//document.getElementById('deleteButton').addEventListener('click', () => excelManager.deleteRow(getRowIndex()));
