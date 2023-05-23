import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { HttpModule } from '@nestjs/axios';
import { TwelveSciService } from './services/twelve-sci.service';
import { TenthService } from './services/tenth.service';

@Module({
  imports: [HttpModule],
  controllers: [AppController],
  providers: [
    AppService,
    TwelveSciService,
    TenthService,
  ],
})
export class AppModule { }
