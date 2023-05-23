import { Body, Controller, Get, Post, Query, Res } from '@nestjs/common';
import { AppService } from './app.service';
import { Response } from 'express';
import { Scrap12SciDto } from './dto/scrap12sci.dto';
import { TwelveSciService } from './services/twelve-sci.service';
import { TenthService } from './services/tenth.service';

@Controller()
export class AppController {
  constructor(
    private readonly appService: AppService,
    private readonly twelevesciService: TwelveSciService,
    private readonly tenthService: TenthService
  ) { }

  @Get()
  getHello(): string {
    return this.appService.getHello();
  }

  @Post('/scrap/12sci')
  scrapResultsfor12sci(@Query() data: Scrap12SciDto, @Res() res: Response) {
    return this.twelevesciService.scrapResultsClass12(data.url, res);
  }

  @Post('/scrap/10th')
  scrapResultsfor10th(@Query() data: Scrap12SciDto, @Res() res: Response) {
    return this.tenthService.scrapResultsClass10(data.url, res);
  }
}
