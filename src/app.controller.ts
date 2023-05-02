import { Body, Controller, Get, Post, Res } from '@nestjs/common';
import { AppService } from './app.service';
import { Response } from 'express';
import { Scrap12SciDto } from './dto/scrap12sci.dto';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) { }

  @Get()
  getHello(): string {
    return this.appService.getHello();
  }

  @Post('/scrap/12sci')
  scrapResultsfor12sci(@Body() data: Scrap12SciDto, @Res() res: Response) {
    return this.appService.scrapResultsClass12(data.url, res);
  }
}
