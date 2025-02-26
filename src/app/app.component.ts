import {Component} from '@angular/core';
import * as XLSX from 'xlsx';
import xlsx from 'node-xlsx';
import {MatIcon} from '@angular/material/icon';
import {DatePipe} from '@angular/common';

@Component({
  selector: 'app-root',
  imports: [
    MatIcon,
    DatePipe
  ],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  host_equiv: { [key: string]: string } = {
    "srv028": "wycom"
  }

  ignore_host_prefixes = {
    "intune": ["accu-", "srv", "sv"],
  }

  kace: any;
  ad: any;
  trend: any;
  tenable: any;
  intune: any;
  blumira: any;

  warnings: string[] = [];
  results: { [key: string]: string[] } = {};
  showResults = false;

  fileChange(target: EventTarget | null) {
    const targetElement = target as HTMLInputElement;
    const file = targetElement.files?.item(0);
    if (file) {
      const fileReader = new FileReader();
      fileReader.onload = (e) => {
        this.warnings = [];
        this.results = {};

        this.kace = this.readExcelTab(fileReader, "Kace");
        this.ad = this.readExcelTab(fileReader, "AD");
        this.trend = this.readExcelTab(fileReader, "Trend");
        this.tenable = this.readExcelTab(fileReader, "Tenable");
        this.intune = this.readExcelTab(fileReader, "Intune");
        this.blumira = this.readExcelTab(fileReader, "Blumira");

        if (this.kace.length == 0) {
          this.warnings.push("No Kace data found");
        }
        if (this.ad.length == 0) {
          this.warnings.push("No AD data found");
        }
        if (this.trend.length == 0) {
          this.warnings.push("No Trend data found");
        }
        if (this.tenable.length == 0) {
          this.warnings.push("No Tenable data found");
        }
        if (this.intune.length == 0) {
          this.warnings.push("No Intune data found");
        }
        if (this.blumira.length == 0) {
          this.warnings.push("No Blumira data found");
        }

        if (this.warnings.length == 0) {
          this.reconcile();
        }

        this.showResults = true;
      };
      fileReader.readAsArrayBuffer(file);
    }
  }

  find_record(data: any, nameCol: string, name: string) {
    // Strip off the domain from the hostname, if there is a domain
    name = name.trim().toLowerCase();
    let idx = name.indexOf(".");
    if (idx != -1)
      name = name.substring(0, idx);

    // Substitute names, if appropriate
    if (this.host_equiv[name])
      name = this.host_equiv[name];

    for (const row of data) {
      let actualName = row[nameCol].trim().toLowerCase();
      let idx = name.indexOf(".");
      if (idx != -1)
        actualName = actualName.substring(0, idx);
      if (this.host_equiv[actualName])
        actualName = this.host_equiv[actualName];

      if (actualName == name)
        return row;
    }

    return undefined;
  }

  evaluate_by_name(name: string, data: any, nameCol: string, ipCol: string = "", skip_prefixes: string[] = []) {
    let evalResults = this.results[name];
    if (!evalResults) {
      evalResults = [];
      this.results[name] = evalResults;
    }

    // Look at each row in Kace
    for (const k_row of this.kace) {
      if (!k_row["Name"])
        continue;
      const kaceName = k_row["Name"].toLowerCase().trim();

      let skip = false;
      for (const prefix of skip_prefixes) {
        if (kaceName.startsWith(prefix)) {
          skip = true;
          break;
        }
      }
      if (skip)
        continue;

      const rec = this.find_record(data, nameCol, kaceName);
      if (rec) {
        if (ipCol && rec[ipCol]) {
          if (k_row["IP Address"] != rec[ipCol] && !rec[nameCol].toLowerCase().includes("lap")) {
            evalResults.push(`${kaceName} exists but has the wrong IP address`);
          }
        }
      } else {
        evalResults.push(`${kaceName} exists in Kace but not ${name}`);
      }
    }

    // Look at each row in the target record set
    for (const d_row of data) {
      const hostname = d_row[nameCol];
      const rec = this.find_record(this.kace, "Name", hostname);
      if (!rec) {
        evalResults.push(`${hostname} exists in ${name} but not Kace`);
      }
    }
  }

  reconcile() {
    this.evaluate_by_name("Active Directory", this.ad, "Name")
    this.evaluate_by_name("Trend", this.trend, "Endpoint")
    this.evaluate_by_name("Tenable", this.tenable, "Host Name", "IPv4 Address")
    this.evaluate_by_name("Blumira", this.blumira, "Device Name", "Device Address")
    this.evaluate_by_name("Intune", this.intune, "Device name", "", this.ignore_host_prefixes["intune"])
  }

  readExcelTab(fr: FileReader, sheet: string) {
    const arrayBuffer: any = fr.result;
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, {type: 'array'});
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);
  }

  goBack() {
    this.showResults = false;
  }

  protected readonly JSON = JSON;
  protected readonly Object = Object;

  downloadExcel() {
    const data = [];

    const date = new Date();
    data.push(["Asset Systems Analysis Results"]);
    data.push([date.toLocaleString()]);
    data.push([]);
    data.push(["System", "Issue"]);
    for (const key of Object.keys(this.results).sort()) {
      const issues = this.results[key];
      for (const issue of issues) {
        data.push([key, issue]);
      }
    }
    const buffer = xlsx.build([{name: 'Asset Analysis Results', data: data, options: {}}]);
    const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'asset_analysis_results.xlsx';
    a.click();
    window.URL.revokeObjectURL(url);
  }

  protected readonly Date = Date;
}
