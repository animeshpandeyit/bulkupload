import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { HttpClient } from '@angular/common/http';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent {
  title = 'saver';
  // Mock API response
  mockApiResponse = [
    { id: 1, isRedFlag: true, ServerMsh: 'Server down' },
    { id: 2, isRedFlag: false, ServerMsh: 'All systems operational' },
    { id: 3, isRedFlag: true, ServerMsh: 'Maintenance required' },
  ];

  onFileChange(event: any) {
    const target: DataTransfer = <DataTransfer>event.target;

    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }

    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const workbook: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      const wsname: string = workbook.SheetNames[0]; // Read the first sheet
      const worksheet: XLSX.WorkSheet = workbook.Sheets[wsname];

      // Convert worksheet data to JSON
      const excelData = XLSX.utils.sheet_to_json(worksheet);

      // Simulate API call with mock data
      this.processApiResponse(excelData, this.mockApiResponse);
    };

    reader.readAsBinaryString(target.files[0]);
  }

  processApiResponse(excelData: any[], apiResponse: any[]) {
    // Modify the Excel data based on the mock API response
    const modifiedData = excelData.map((row: any) => {
      const apiMatch = apiResponse.find((item: any) => item.id === row.id);

      if (apiMatch) {
        return {
          ...row,
          Status: apiMatch.isRedFlag
            ? `Error: ${apiMatch.ServerMsh}`
            : 'All Good',
        };
      } else {
        return { ...row, Status: 'No Data Available' };
      }
    });

    // Convert modified data back to worksheet
    const newWorksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(modifiedData);
    const newWorkbook: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'UpdatedSheet');

    // Download the modified Excel file
    const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    saveAs(
      new Blob([wbout], { type: 'application/octet-stream' }),
      'Updated_File.xlsx'
    );
  }
}
