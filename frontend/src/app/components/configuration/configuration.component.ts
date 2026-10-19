import { Component, EventEmitter, Input, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormArray, FormControl, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { MatButtonModule } from '@angular/material/button';
import { MatTabsModule } from '@angular/material/tabs';
import { Configuration } from '../../models';
const text = (v: string) =>
  new FormControl(v, {
    nonNullable: true,
    validators: [Validators.required, Validators.maxLength(200)],
  });
const lines = (v: string) =>
  v
    .split('\n')
    .map((x) => x.trim())
    .filter(Boolean);
@Component({
  selector: 'app-configuration',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule, MatButtonModule, MatTabsModule],
  templateUrl: './configuration.component.html',
})
export class ConfigurationComponent {
  @Input() set initialConfiguration(c: Configuration | undefined) {
    if (c) this.set(c);
  }
  @Input() busy = false;
  @Output() configurationChange = new EventEmitter<{
    configuration: Configuration;
    valid: boolean;
  }>();
  @Output() importFile = new EventEmitter<File>();
  @Output() exportSettings = new EventEmitter<void>();
  @Output() resetSettings = new EventEmitter<void>();
  readonly form = new FormGroup({
    methods: new FormArray<FormGroup>([]),
    aliases: new FormArray<FormGroup>([]),
    patterns: new FormArray<FormGroup>([]),
    banks: new FormControl('', { nonNullable: true }),
    branches: new FormControl('', { nonNullable: true }),
    nonCard: new FormControl('', { nonNullable: true }),
  });
  readonly search = new FormControl('', { nonNullable: true });
  readonly query = signal('');
  constructor() {
    this.form.valueChanges.subscribe(() => this.emit());
    this.search.valueChanges.subscribe((v) => this.query.set(v.toLowerCase()));
  }
  get methods() {
    return this.form.controls.methods;
  }
  get aliases() {
    return this.form.controls.aliases;
  }
  get patterns() {
    return this.form.controls.patterns;
  }
  get bankNames() {
    return lines(this.form.controls.banks.value);
  }
  get filtered() {
    return this.methods.controls
      .map((group, index) => ({ group, index }))
      .filter(({ group }) =>
        `${group.value.cardName} ${group.value.bank}`.toLowerCase().includes(this.query()),
      );
  }
  private method(cardName = '', bank = '', rate = '0.00') {
    return new FormGroup({
      cardName: text(cardName),
      bank: text(bank),
      rate: new FormControl(rate, {
        nonNullable: true,
        validators: [
          Validators.required,
          Validators.pattern(/^-?\d{1,3}(\.\d{1,6})?$/),
          Validators.min(-100),
          Validators.max(100),
        ],
      }),
    });
  }
  set(c: Configuration) {
    this.methods.clear({ emitEvent: false });
    this.aliases.clear({ emitEvent: false });
    this.patterns.clear({ emitEvent: false });
    c.paymentMethods.forEach((m) =>
      this.methods.push(this.method(m.cardName, m.bank, m.rate), { emitEvent: false }),
    );
    c.aliases.forEach((a) =>
      this.aliases.push(new FormGroup({ alias: text(a.alias), cardName: text(a.cardName) }), {
        emitEvent: false,
      }),
    );
    c.patterns.forEach((p) =>
      this.patterns.push(new FormGroup({ pattern: text(p.pattern), cardName: text(p.cardName) }), {
        emitEvent: false,
      }),
    );
    this.form.patchValue(
      {
        banks: c.banks.join('\n'),
        branches: c.branches.join('\n'),
        nonCard: c.nonCardMethods.join('\n'),
      },
      { emitEvent: false },
    );
    this.emit();
  }
  addMethod() {
    this.search.setValue('');
    this.methods.push(this.method('', this.bankNames[0] || ''));
  }
  addAlias() {
    this.aliases.push(new FormGroup({ alias: text(''), cardName: text('') }));
  }
  addPattern() {
    this.patterns.push(new FormGroup({ pattern: text(''), cardName: text('') }));
  }
  emit() {
    const v = this.form.getRawValue();
    this.configurationChange.emit({
      configuration: {
        paymentMethods: v.methods.map((m) => ({
          cardName: String(m['cardName'] ?? ''),
          bank: String(m['bank'] ?? ''),
          rate: String(m['rate'] ?? ''),
        })),
        aliases: v.aliases as Configuration['aliases'],
        patterns: v.patterns as Configuration['patterns'],
        banks: lines(v.banks),
        branches: lines(v.branches),
        nonCardMethods: lines(v.nonCard),
      },
      valid: this.form.valid,
    });
  }
  pick(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) this.importFile.emit(file);
  }
}
