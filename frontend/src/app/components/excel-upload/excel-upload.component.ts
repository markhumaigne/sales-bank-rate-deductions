import { Component, EventEmitter, Output, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormControl, ReactiveFormsModule } from '@angular/forms';
import { MatButtonModule } from '@angular/material/button';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { ExcelApiService, readableError } from '../../services/excel-api.service';
import { ReportUpload, SheetPreview, WorkbookPreview, CostBasis } from '../../models';

@Component({
  selector: 'app-excel-upload',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, MatButtonModule, MatProgressBarModule],
  templateUrl: './excel-upload.component.html',
})
export class ExcelUploadComponent {
  @Output() reportChange = new EventEmitter<ReportUpload | undefined>();
  readonly api = inject(ExcelApiService);
  readonly loading = signal(false);
  readonly error = signal('');
  workbook = signal<WorkbookPreview | undefined>(undefined);
  selected = signal<SheetPreview | undefined>(undefined);
  file?: File;
  controls: Record<string, FormControl<string>> = {};
  readonly sheetControl = new FormControl('', { nonNullable: true });
  readonly headerControl = new FormControl(1, { nonNullable: true });
  readonly costBasisControl = new FormControl<CostBasis>('auto', { nonNullable: true });
  readonly paymentScopeControl = new FormControl<'invoice' | 'line'>('invoice', {
    nonNullable: true,
  });
  constructor() {
    this.costBasisControl.valueChanges.subscribe(() => this.emit());
    this.paymentScopeControl.valueChanges.subscribe(() => this.emit());
  }
  private generation = 0;
  get title() {
    return 'Sales Report';
  }
  get fields() {
    return [
      'productId',
      'date',
      'branch',
      'invoice',
      'product',
      'quantity',
      'price',
      'cost',
      'discount',
      'lineTotal',
      'payment',
      'paymentAmount',
    ];
  }
  get required() {
    return ['invoice', 'product', 'quantity', 'lineTotal'];
  }
  labels: Record<string, string> = {
    productId: 'Product ID / SKU',
    price: 'Price per unit',
    cost: 'Product cost',
    discount: 'Discount',
    lineTotal: 'Final net line total',
    paymentAmount: 'Explicit payment amount (optional)',
    date: 'Transaction date (MM/DD/YYYY or ISO)',
    invoice: 'Invoice reference',
    branch: 'Branch',
    product: 'Product name',
    quantity: 'Quantity',
    payment: 'Payment details',
  };
  get missing() {
    return this.required.some((k) => !this.controls[k]?.value) || this.ambiguousMappings.length > 0;
  }
  get ambiguousMappings() {
    return (this.selected()?.ambiguousFields ?? [])
      .filter((key) => !this.controls[key]?.value)
      .map((key) => this.labels[key]);
  }
  pick(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) void this.load(file);
  }
  drop(event: DragEvent) {
    event.preventDefault();
    const file = event.dataTransfer?.files[0];
    if (file) void this.load(file);
  }
  async load(file: File, headerRow?: number) {
    const generation = ++this.generation;
    this.reportChange.emit(undefined);
    this.error.set('');
    this.loading.set(true);
    this.file = file;
    this.selected.set(undefined);
    if (!/\.xlsx?$/i.test(file.name) || file.size > 10 * 1024 * 1024) {
      this.error.set('Choose an .xlsx or .xls workbook no larger than 10 MB.');
      this.loading.set(false);
      this.workbook.set(undefined);
      return;
    }
    try {
      const workbook = await this.api.preview(file, headerRow);
      if (generation !== this.generation) return;
      this.workbook.set(workbook);
      const sheets = [...workbook.sheets].sort(
        (a, b) =>
          this.required.filter((k) => b.suggestedMapping[k]).length -
          this.required.filter((k) => a.suggestedMapping[k]).length,
      );
      this.selectSheet(
        sheets.find((s) => s.name === this.sheetControl.value)?.name || sheets[0].name,
      );
    } catch (e) {
      if (generation === this.generation) {
        this.error.set(await readableError(e));
        this.workbook.set(undefined);
      }
    } finally {
      if (generation === this.generation) this.loading.set(false);
    }
  }
  selectSheet(name: string) {
    const sheet = this.workbook()?.sheets.find((s) => s.name === name);
    if (!sheet) return;
    this.selected.set(sheet);
    this.sheetControl.setValue(name);
    this.headerControl.setValue(sheet.headerRow);
    this.controls = {};
    for (const field of this.fields) {
      const c = new FormControl(sheet.suggestedMapping[field] || '', { nonNullable: true });
      c.valueChanges.subscribe(() => this.emit());
      this.controls[field] = c;
    }
    this.emit();
  }
  emit() {
    const sheet = this.selected();
    if (!this.file || !sheet) return;
    const mapping = Object.fromEntries(this.fields.map((k) => [k, this.controls[k]?.value || '']));
    this.reportChange.emit({
      file: this.file,
      selection: {
        sheet: sheet.name,
        headerRow: sheet.headerRow,
        mapping,
        costBasis: this.costBasisControl.value,
        paymentScope: this.paymentScopeControl.value,
      },
      preview: sheet,
      valid: !this.missing && sheet.recordCount > 0,
    });
  }
  reread() {
    if (this.file) void this.load(this.file, this.headerControl.value);
  }
}
