import {Component} from '@angular/core';
import {MatIcon} from '@angular/material/icon';
import {DatePipe} from '@angular/common';
import * as Excel from 'exceljs';

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
  workbook!: Excel.Workbook;

  warnings: string[] = [];
  results: { [key: string]: string[] } = {};
  showResults = false;

  fileChange(target: EventTarget | null) {
    const targetElement = target as HTMLInputElement;
    const file = targetElement.files?.item(0);
    if (file) {
      const fileReader = new FileReader();
      fileReader.onload = async (e) => {
        this.warnings = [];
        this.results = {};

        const arrayBuffer: any = fileReader.result;
        const workbook = new Excel.Workbook();
        this.workbook = await workbook.xlsx.load(arrayBuffer);

        this.kace = this.readExcelTab("Kace");
        this.ad = this.readExcelTab("AD");
        this.trend = this.readExcelTab("Trend");
        this.tenable = this.readExcelTab("Tenable");
        this.intune = this.readExcelTab("Intune");
        this.blumira = this.readExcelTab("Blumira");

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

  find_record(data: any, nameHeader: string, name: string) {
    // Strip off the domain from the hostname, if there is a domain
    name = this.host(name.trim().toLowerCase());

    // Substitute names, if appropriate
    if (this.host_equiv[name])
      name = this.host_equiv[name];

    let nameColIdx = this.get_column_index(nameHeader, data);
    for (let i = 1; i < data.length; ++i) {
      const row = data[i];
      let actualName = this.host(row[nameColIdx].trim().toLowerCase());
      if (this.host_equiv[actualName])
        actualName = this.host_equiv[actualName];

      if (actualName == name)
        return row;
    }
    return undefined;
  }

  get_column_index(name: string, data: any[]): number {
    let row = data[0];
    if (Array.isArray(row)) {
      row = data[0];
    }
    return row.indexOf(name);
  }

  evaluate_by_name(name: string, data: any, nameHeader: string, ipHeader: string = "", skip_prefixes: string[] = []) {
    let evalResults = this.results[name];
    if (!evalResults) {
      evalResults = [];
      this.results[name] = evalResults;
    }

    // Look at each row in Kace
    const kaceNameIdx = this.get_column_index("Name", this.kace);
    const kaceIpIdx = this.get_column_index("IP Address", this.kace);
    const dataNameIdx = this.get_column_index(nameHeader, data);
    const dataIpIdx = this.get_column_index(ipHeader, data);
    for (let i = 1; i < this.kace.length; ++i) {
      const k_row = this.kace[i];
      const kaceName = k_row[kaceNameIdx].toLowerCase().trim();

      let skip = false;
      for (const prefix of skip_prefixes) {
        if (kaceName.startsWith(prefix)) {
          skip = true;
          break;
        }
      }
      if (skip)
        continue;

      const rec = this.find_record(data, nameHeader, kaceName);
      if (rec) {
        if (ipHeader && kaceIpIdx != -1 && rec[dataIpIdx] && k_row[kaceIpIdx] != rec[dataIpIdx] && !rec[dataNameIdx].toLowerCase().includes("lap")) {
          evalResults.push(`${kaceName} exists but has the wrong IP address`);
        }
      } else {
        evalResults.push(`${kaceName} exists in Kace but not ${name}`);
      }
    }

    // Look at each row in the target record set
    for (let idx = 1; idx < data.length; ++idx) {
      const d_row = data[idx];
      const hostname = d_row[dataNameIdx].toLowerCase().trim();
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

  readExcelTab(sheet: string): any[] {
    const worksheet = this.workbook.getWorksheet(sheet)!;
    const ret: any[] = [];
    worksheet.eachRow((row, rowNumber) => {
      const r = row.values as any[];
      r.shift(); // First cell is always empty
      ret.push(r);
    });
    return ret;
  }

  goBack() {
    this.showResults = false;
  }

  protected readonly JSON = JSON;
  protected readonly Object = Object;

  async downloadExcel() {
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

    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet("Asset Analysis Results");
    sheet.addRows(data);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'asset_analysis_results.xlsx';
    a.click();
    window.URL.revokeObjectURL(url);
  }

  host(fqdn: string): string {
    const idx = fqdn.indexOf(".");
    if (idx == -1)
      return fqdn;
    return fqdn.substring(0, idx);
  }

  protected readonly Date = Date;
}
