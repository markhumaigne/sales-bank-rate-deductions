import { Component, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MatButtonModule } from '@angular/material/button';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { ExcelUploadComponent } from '../excel-upload/excel-upload.component';
import { ConfigurationComponent } from '../configuration/configuration.component';
import { CalculationResultsComponent } from '../calculation-results/calculation-results.component';
import { CalculationResponse, Configuration, ReportUpload } from '../../models';
import { ExcelApiService, readableError, saveDownload } from '../../services/excel-api.service';
@Component({
  selector: 'app-root',
  standalone: true,
  imports: [
    CommonModule,
    MatButtonModule,
    MatProgressBarModule,
    ExcelUploadComponent,
    ConfigurationComponent,
    CalculationResultsComponent,
  ],
  templateUrl: './dashboard.component.html',
})
export class DashboardComponent {
  readonly api = inject(ExcelApiService);
  readonly configSeed = signal<Configuration | undefined>(undefined);
  readonly response = signal<CalculationResponse | undefined>(undefined);
  readonly busy = signal(false);
  readonly configBusy = signal(false);
  readonly downloading = signal(false);
  readonly error = signal('');
  readonly notice = signal('');
  readonly loadingDefaults = signal(true);
  sales?: ReportUpload;
  configuration?: Configuration;
  configValid = false;
  revision = 0;
  constructor() {
    void this.loadDefaults();
  }
  async loadDefaults() {
    this.loadingDefaults.set(true);
    this.error.set('');
    try {
      const c = await this.api.defaults();
      this.invalidate();
      this.configSeed.set(c);
      this.notice.set('Default settings loaded.');
    } catch (e) {
      this.error.set(await readableError(e));
    } finally {
      this.loadingDefaults.set(false);
    }
  }
  setReport(report: ReportUpload | undefined) {
    this.sales = report;
    this.invalidate();
  }
  setConfig(event: { configuration: Configuration; valid: boolean }) {
    this.configuration = event.configuration;
    this.configValid = event.valid;
    this.invalidate();
  }
  invalidate() {
    this.revision++;
    if (this.response()) {
      this.response.set(undefined);
      this.notice.set('Inputs changed. Calculate again to update the results.');
    }
  }
  get ready() {
    return (
      !!this.sales?.valid &&
      this.configValid &&
      !!this.configuration &&
      !this.busy() &&
      !this.loadingDefaults() &&
      !this.configBusy()
    );
  }
  async calculate() {
    if (!this.ready) return;
    const revision = this.revision;
    this.busy.set(true);
    this.error.set('');
    this.notice.set('');
    this.response.set(undefined);
    try {
      const response = await this.api.calculate(this.sales!, this.configuration!);
      if (revision === this.revision) {
        this.response.set(response);
        this.notice.set('Calculation complete. Your Excel report is ready.');
        setTimeout(
          () =>
            document
              .getElementById('results')
              ?.scrollIntoView({ behavior: 'smooth', block: 'start' }),
          100,
        );
      } else
        this.notice.set(
          'Inputs changed during processing. Calculate again with the current settings.',
        );
    } catch (e) {
      this.error.set(await readableError(e));
    } finally {
      this.busy.set(false);
    }
  }
  async importConfig(file: File) {
    this.configBusy.set(true);
    this.error.set('');
    try {
      const c = await this.api.import(file);
      this.invalidate();
      this.configSeed.set(c);
      this.notice.set('Configuration imported. Review the settings before calculating.');
    } catch (e) {
      this.error.set(await readableError(e));
    } finally {
      this.configBusy.set(false);
    }
  }
  async exportConfig() {
    if (!this.configuration) return;
    this.configBusy.set(true);
    this.error.set('');
    try {
      saveDownload(await this.api.exportConfig(this.configuration), 'Bank_Rate_Configuration.xlsx');
    } catch (e) {
      this.error.set(await readableError(e));
    } finally {
      this.configBusy.set(false);
    }
  }
  async download() {
    const response = this.response();
    if (!response) return;
    this.downloading.set(true);
    this.error.set('');
    try {
      saveDownload(await this.api.exportReport(response), 'Sales_Bank_Rate_Deductions.xlsx');
    } catch (e) {
      this.error.set(await readableError(e));
    } finally {
      this.downloading.set(false);
    }
  }
}
