import { BadRequestException, Injectable } from '@nestjs/common';
import { z } from 'zod';
import Decimal from 'decimal.js';
import defaults from './bank-rates.json';
import branches from './branches.json';
import { Configuration } from '../../../shared/models';

export const normalize = (value: string) => value.replace(/\s+/g, ' ').trim().toUpperCase();
const name = z.string().trim().min(1).max(200);
const schema = z
  .object({
    paymentMethods: z
      .array(
        z.object({
          cardName: name,
          bank: name,
          rate: z
            .string()
            .regex(/^-?\d{1,3}(\.\d{1,6})?$/, 'Rate must be a percentage from -100 to 100.'),
          branchRates: z
            .record(
              z
                .string()
                .min(1)
                .max(200)
                .refine((value) => !!value.trim()),
              z.string().regex(/^-?\d{1,3}(\.\d{1,6})?$/),
            )
            .optional(),
        }),
      )
      .max(1000),
    banks: z.array(name).max(500),
    branches: z.array(name).max(1000),
    aliases: z.array(z.object({ alias: name, cardName: name })).max(1000),
    patterns: z.array(z.object({ pattern: name, cardName: name })).max(200),
    nonCardMethods: z.array(name).max(200),
  })
  .strict();

@Injectable()
export class ConfigService {
  defaults(): Configuration {
    return structuredClone({ ...defaults, branches });
  }
  validate(input: unknown): Configuration {
    const parsed = schema.safeParse(input);
    if (!parsed.success)
      throw new BadRequestException(
        'Invalid configuration: ' +
          parsed.error.issues
            .map((i) => `${i.path.join('.')}: ${i.message}`)
            .slice(0, 8)
            .join('; '),
      );
    const c = parsed.data;
    const fail = (message: string): never => {
      throw new BadRequestException(message);
    };
    const unique = (values: string[], label: string) => {
      if (new Set(values.map(normalize)).size !== values.length)
        fail(`Duplicate ${label}. Names are case-insensitive.`);
    };
    unique(c.banks, 'bank');
    unique(c.branches, 'branch');
    unique(
      c.paymentMethods.map((m) => m.cardName),
      'card name',
    );
    unique(
      c.aliases.map((a) => a.alias),
      'alias',
    );
    unique(
      c.patterns.map((p) => p.pattern),
      'pattern',
    );
    unique(c.nonCardMethods, 'non-card method');
    const cards = new Set(c.paymentMethods.map((m) => normalize(m.cardName)));
    for (const m of c.paymentMethods) {
      if (!c.banks.some((b) => normalize(b) === normalize(m.bank)))
        fail(`Bank "${m.bank}" must be in the bank reference list.`);
      if (new Decimal(m.rate).abs().gt(100))
        fail(`Rate for ${m.cardName} must be between -100 and 100.`);
      unique(Object.keys(m.branchRates ?? {}), 'branch rate');
      for (const [branch, rate] of Object.entries(m.branchRates ?? {})) {
        if (!c.branches.some((b) => normalize(b) === normalize(branch)))
          fail(`Branch "${branch}" for ${m.cardName} must be in the branch reference list.`);
        if (new Decimal(rate).abs().gt(100))
          fail(`Branch rate for ${m.cardName} must be between -100 and 100.`);
      }
    }
    for (const a of c.aliases) {
      if (!cards.has(normalize(a.cardName)))
        fail(`Alias "${a.alias}" refers to an unknown card name.`);
      if (cards.has(normalize(a.alias))) fail(`Alias "${a.alias}" conflicts with a card name.`);
    }
    for (const p of c.patterns) {
      if (!cards.has(normalize(p.cardName)))
        fail(`Pattern "${p.pattern}" refers to an unknown card name.`);
      if (!p.pattern.replace(/\*/g, '').trim())
        fail('Patterns must contain literal text; a catch-all pattern is not allowed.');
    }
    for (const n of c.nonCardMethods)
      if (cards.has(normalize(n)) || c.aliases.some((a) => normalize(a.alias) === normalize(n)))
        fail(`Non-card method "${n}" conflicts with a configured card or alias.`);
    return c;
  }
}
