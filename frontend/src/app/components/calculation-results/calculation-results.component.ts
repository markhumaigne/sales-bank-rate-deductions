import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormControl, ReactiveFormsModule } from '@angular/forms';
import { MatButtonModule } from '@angular/material/button';
import { MatTabsModule } from '@angular/material/tabs';
import {
  Calculation,
  ITEM_COLUMNS,
  BRANCH_DATE_SUMMARY_COLUMNS,
  BRANCH_DATE_SUMMARY_SHEET,
  BRANCH_DATE_SUMMARY_TITLE,
  showInItemBankDeductions,
} from '../../models';
type Column = { header: string; key: string; format: string };
@Component({
  selector: 'app-calculation-results',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, MatButtonModule, MatTabsModule],
  templateUrl: './calculation-results.component.html',
})
export class CalculationResultsComponent {
  @Input({ required: true }) set calculation(value: Calculation) {
    this.result = value;
    this.page = 0;
    this.data = [
      value.items.filter(showInItemBankDeductions).map((i) => ({ ...i })),
      value.branchDateSummary.map((b) => ({ ...b })),
      value.reviews.map((r) => ({ ...r })),
    ];
  }
  @Input() downloading = false;
  @Output() downloadReport = new EventEmitter<void>();
  result!: Calculation;
  view = 0;
  page = 0;
  readonly pageSize = 50;
  readonly branchDateSummaryTitle = BRANCH_DATE_SUMMARY_TITLE;
  readonly branchDateSummarySheet = BRANCH_DATE_SUMMARY_SHEET;
  sortKey = 'sourceRow';
  ascending = true;
  readonly search = new FormControl('', { nonNullable: true });
  data: Record<string, unknown>[][] = [];
  constructor() {
    this.search.valueChanges.subscribe(() => (this.page = 0));
  }
  readonly columnSets: Column[][] = [
    ITEM_COLUMNS.map((c) => ({ ...c })),
    BRANCH_DATE_SUMMARY_COLUMNS.map((c) => ({ ...c })),
    [
      { header: 'Product ID', key: 'productId', format: 'text' },
      { header: 'Sales Date', key: 'salesDate', format: 'date' },
      { header: 'Branch', key: 'branch', format: 'text' },
      { header: 'Invoice Number', key: 'invoice', format: 'text' },
      { header: 'Product Name', key: 'product', format: 'text' },
      { header: 'Sales Source Row', key: 'sourceRow', format: 'number' },
      { header: 'Status', key: 'status', format: 'text' },
      { header: 'Reason / Warning', key: 'reason', format: 'text' },
    ],
  ];
  get columns() {
    return this.columnSets[this.view];
  }
  get matchedItemCount() {
    return this.data[0]?.length ?? 0;
  }
  get filtered() {
    const query = this.search.value.toLowerCase();
    const values = (this.data[this.view] || []).filter(
      (r) =>
        !query ||
        this.columns.some((c) =>
          String(r[c.key] ?? '')
            .toLowerCase()
            .includes(query),
        ),
    );
    const col = this.columns.find((c) => c.key === this.sortKey);
    return values.sort((a, b) => {
      const av = a[this.sortKey],
        bv = b[this.sortKey];
      let order = 0;
      if (
        col &&
        ['money', 'number', 'percent'].includes(col.format) &&
        this.numeric(av) !== null &&
        this.numeric(bv) !== null
      )
        order = Number(av) - Number(bv);
      else order = String(av ?? '').localeCompare(String(bv ?? ''));
      return (
        (this.ascending ? order : -order) ||
        (this.view === 1 ? String(a['branch']).localeCompare(String(b['branch'])) : 0) ||
        Number(a['sourceRow'] || 0) - Number(b['sourceRow'] || 0)
      );
    });
  }
  get visibleRows() {
    return this.filtered.slice(this.page * this.pageSize, (this.page + 1) * this.pageSize);
  }
  get pages() {
    return Math.max(1, Math.ceil(this.filtered.length / this.pageSize));
  }
  changeView(index: number) {
    this.view = index;
    this.page = 0;
    this.sortKey = index === 1 ? 'date' : 'sourceRow';
    this.ascending = true;
    this.search.setValue('');
  }
  sort(key: string) {
    if (this.sortKey === key) this.ascending = !this.ascending;
    else {
      this.sortKey = key;
      this.ascending = true;
    }
    this.page = 0;
  }
  display(v: unknown) {
    return v === null || v === undefined || v === '' ? '—' : String(v);
  }
  numeric(v: unknown) {
    const text = String(v ?? '');
    return /^-?\d+(\.\d+)?$/.test(text) ? text : null;
  }
  calendarDate(v: unknown) {
    const d = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v ?? ''));
    return d
      ? d[3] +
          ' ' +
          ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][
            Number(d[2]) - 1
          ] +
          ' ' +
          d[1]
      : '—';
  }
  warning(row: Record<string, unknown>) {
    return Array.isArray(row['warnings']) ? (row['warnings'] as string[]).join(' ') : '';
  }
  needsReview(row: Record<string, unknown>) {
    return (
      this.view === 2 ||
      (Array.isArray(row['issues']) && row['issues'].length > 0) ||
      (Array.isArray(row['warnings']) && row['warnings'].length > 0)
    );
  }
}
