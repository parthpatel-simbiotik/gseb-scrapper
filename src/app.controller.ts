import { Body, Controller, Get, Post, Query, Res } from '@nestjs/common';
import { AppService } from './app.service';
import { Response } from 'express';
import { Scrap12SciDto } from './dto/scrap12sci.dto';
import { TwelveSciService } from './services/twelve-sci.service';
import { TenthService } from './services/tenth.service';
import { TwelveGenService } from './services/twelve-gen.service';

@Controller()
export class AppController {
  constructor(
    private readonly appService: AppService,
    private readonly tweleveSciService: TwelveSciService,
    private readonly tweleveGenService: TwelveGenService,
    private readonly tenthService: TenthService
  ) { }

  @Get()
  getHello(): string {
    return this.appService.getHello();
  }

  @Post('/scrap/12gen')
  scrapResultsfor12gen(@Res() res: Response, @Query('filename') filename: string) {
    return this.tweleveGenService.scrapRecurrsively(filename, res);
  }

  @Post('/scrap/12sci')
  scrapResultsfor12sci(@Res() res: Response) {
    return this.tweleveSciService.scrapRecurrsively(res);
  }

  @Post('/scrap/10th')
  scrapResultsfor10th(@Res() res: Response) {
    return this.tenthService.scrapRecurrsively(res);
  }
}
