import {
  BadRequestException,
  Body,
  Controller,
  Get,
  Post,
  Query,
  Res,
  UploadedFile,
  UseInterceptors,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { createHmac, randomBytes, timingSafeEqual } from 'crypto';
import { z } from 'zod';
import { Calculation, Selection } from '../../../shared/models';
import { CalculationService } from '../calculation/calculation.service';
import { ConfigService } from '../config/config.service';
import { ExcelService } from './excel.service';
import { ExcelExportService } from './excel-export.service';
import { uploadOptions } from './upload';

const selectionSchema = z
  .object({
    sheet: z.string().min(1).max(200),
    headerRow: z.number().int().min(1).max(50),
    mapping: z.record(z.string().max(200)).refine((v) => Object.keys(v).length <= 12),
    costBasis: z.enum(['auto', 'unit', 'line']).optional(),
    paymentScope: z.enum(['invoice', 'line']).optional(),
  })
  .strict();
@Controller('api')
export class ExcelController {
  private readonly secret = randomBytes(32);
  constructor(
    private readonly excel: ExcelService,
    private readonly calculator: CalculationService,
    private readonly exporter: ExcelExportService,
    private readonly config: ConfigService,
  ) {}
  @Get('config/defaults') defaults() {
    return this.config.defaults();
  }
  @Post('excel/preview')
  @UseInterceptors(FileInterceptor('file', uploadOptions))
  preview(@UploadedFile() file: Express.Multer.File, @Query('headerRow') headerRow?: string) {
    return this.excel.preview(file, headerRow ? Number(headerRow) : undefined);
  }
  @Post('config/import')
  @UseInterceptors(FileInterceptor('file', uploadOptions))
  import(@UploadedFile() file: Express.Multer.File) {
    return this.exporter.import(file);
  }
  @Post('config/export') async configExport(@Body() body: unknown, @Res() response: Response) {
    const c = this.config.validate(body);
    this.download(response, await this.exporter.configuration(c), 'Bank_Rate_Configuration.xlsx');
  }
  @Post('excel/calculate')
  @UseInterceptors(FileInterceptor('sales', uploadOptions))
  calculate(@UploadedFile() file: Express.Multer.File, @Body('options') options: string) {
    let input: { sales: Selection; configuration?: unknown };
    try {
      input = z
        .object({ sales: selectionSchema, configuration: z.unknown() })
        .strict()
        .parse(JSON.parse(options));
    } catch {
      throw new BadRequestException(
        'Provide valid worksheet selections, column mappings, and configuration.',
      );
    }
    const c = this.config.validate(input.configuration);
    const sales = this.excel.parse(file, input.sales);
    const result = this.calculator.calculate(sales, c);
    if (Buffer.byteLength(JSON.stringify(result)) > 24 * 1024 * 1024)
      throw new BadRequestException(
        'The calculated report exceeds the 24 MB session limit. Split the reports into smaller batches.',
      );
    return { result, exportToken: this.sign(result) };
  }
  @Post('excel/export') async export(
    @Body() body: { result?: Calculation; exportToken?: string },
    @Res() response: Response,
  ) {
    if (
      !body?.result ||
      typeof body.exportToken !== 'string' ||
      !/^[a-f0-9]{64}$/.test(body.exportToken)
    )
      throw new BadRequestException('Calculate successfully before downloading a report.');
    const expected = Buffer.from(this.sign(body.result), 'hex'),
      received = Buffer.from(body.exportToken, 'hex');
    if (!timingSafeEqual(expected, received))
      throw new BadRequestException(
        'Calculation has changed or the server restarted. Calculate again before downloading.',
      );
    const date = new Intl.DateTimeFormat('sv-SE', {
      timeZone: 'Asia/Manila',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
    })
      .format(new Date())
      .replace(/[-:]/g, '')
      .replace(' ', '_');
    this.download(
      response,
      await this.exporter.report(body.result),
      `Sales_Bank_Rate_Deductions_${date}.xlsx`,
    );
  }
  private sign(result: Calculation) {
    return createHmac('sha256', this.secret).update(JSON.stringify(result)).digest('hex');
  }
  private download(response: Response, buffer: Buffer, name: string) {
    response.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    response.setHeader('Content-Disposition', `attachment; filename="${name}"`);
    response.setHeader('Cache-Control', 'no-store');
    response.send(buffer);
  }
}
