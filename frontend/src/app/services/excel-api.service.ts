import { Injectable, inject } from '@angular/core';
import { HttpClient, HttpErrorResponse, HttpResponse } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { CalculationResponse, Configuration, ReportUpload, WorkbookPreview } from '../models';
@Injectable({ providedIn: 'root' })
export class ExcelApiService {
  private readonly http = inject(HttpClient);
  defaults() {
    return firstValueFrom(this.http.get<Configuration>('/api/config/defaults'));
  }
  preview(file: File, headerRow?: number) {
    const data = new FormData();
    data.append('file', file);
    return firstValueFrom(
      this.http.post<WorkbookPreview>(
        `/api/excel/preview${headerRow ? '?headerRow=' + headerRow : ''}`,
        data,
      ),
    );
  }
  calculate(sales: ReportUpload, configuration: Configuration) {
    const data = new FormData();
    data.append('sales', sales.file);
    data.append('options', JSON.stringify({ sales: sales.selection, configuration }));
    return firstValueFrom(this.http.post<CalculationResponse>('/api/excel/calculate', data));
  }
  import(file: File) {
    const data = new FormData();
    data.append('file', file);
    return firstValueFrom(this.http.post<Configuration>('/api/config/import', data));
  }
  exportConfig(configuration: Configuration) {
    return firstValueFrom(
      this.http.post('/api/config/export', configuration, {
        responseType: 'blob',
        observe: 'response',
      }),
    );
  }
  exportReport(calculation: CalculationResponse) {
    return firstValueFrom(
      this.http.post('/api/excel/export', calculation, {
        responseType: 'blob',
        observe: 'response',
      }),
    );
  }
}
export async function readableError(error: unknown): Promise<string> {
  if (error instanceof HttpErrorResponse) {
    if (error.status === 0)
      return 'Cannot reach the local API. Start the NestJS backend on port 3000 and try again.';
    let value = error.error;
    if (value instanceof Blob) {
      try {
        value = JSON.parse(await value.text());
      } catch {
        return 'The Excel download failed. Please calculate again.';
      }
    }
    return Array.isArray(value?.message)
      ? value.message.join('; ')
      : value?.message || error.message;
  }
  return error instanceof Error ? error.message : 'An unexpected error occurred. Please try again.';
}
export function saveDownload(response: HttpResponse<Blob>, fallback: string) {
  if (!response.body) throw new Error('The server returned an empty download.');
  const name =
    /filename="([^"]+)"/.exec(response.headers.get('Content-Disposition') || '')?.[1] || fallback;
  const url = URL.createObjectURL(response.body),
    link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
