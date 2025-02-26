import {Component} from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  imports: [],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  host_equiv = {
    "srv028": "wycom"
  }

  ignore_host_prefixes = {
    "intune": ["accu-", "srv", "sv"],
  }

  fileChange(target: EventTarget | null) {
    const targetElement = target as HTMLInputElement;
    const file = targetElement.files?.item(0);
    if (file) {
      let excelData: any;
      const fileReader = new FileReader();
      fileReader.onload = (e) => {
        const arrayBuffer: any = fileReader.result;
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, {type: 'array'});
        excelData = XLSX.utils.sheet_to_json(workbook.Sheets["Kace"]);
        console.log(excelData);
      };
      fileReader.readAsArrayBuffer(file);
    }
  }
}
