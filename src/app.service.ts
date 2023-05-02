import { HttpService } from '@nestjs/axios';
import { Injectable } from '@nestjs/common';
import { Workbook } from 'exceljs';
import { Response } from 'express';
import { existsSync, mkdir, mkdirSync, readFileSync, writeFileSync } from 'fs';
import * as moment from 'moment';
import { firstValueFrom } from 'rxjs';
const fs = require('fs');
const cheerio = require('cheerio');
const csv = require('csv-parser');


@Injectable()
export class AppService {
  constructor(private httpService: HttpService) { }

  getHello(): string {
    return 'Hello World!';
  }

  async scrapResultsClass12(baseurl: string, res: Response) {
    var inputFilePath = 'files/input/gm1.csv';
    var outputFileDirectory = `files/output${inputFilePath.substring(inputFilePath.lastIndexOf('/'), inputFilePath.lastIndexOf('.'))}/`;
    console.log('HAELLOEAO', inputFilePath, outputFileDirectory);

    var seatNumbers = await this.csvToRows(inputFilePath);
    console.log('seatNumbers', seatNumbers);

    const workbook = new Workbook();
    workbook.creator = 'Bhagwan';
    workbook.created = new Date();
    const sheet = workbook.addWorksheet('HSC-ScienceResults');
    sheet.columns = [
      { header: 'SeatNumber', key: 'SeatNumber' },
      { header: 'Name', key: 'Name' },
      { header: 'ScienceMarks', key: 'ScienceMarks' },
      { header: 'MathsMarks', key: 'MathsMarks' },
      { header: 'EnglishMarks', key: 'EnglishMarks' },
      { header: 'GujaratiMarks', key: 'GujaratiMarks' },
      { header: 'SanskritMarks', key: 'SanskritMarks' },
    ];


    // todo: api call
    for (let i = 0; i < seatNumbers.length; i++) {
      console.log(`[Scrapping from ${baseurl}] [${i}/${seatNumbers.length}]`);
      const seatNum = seatNumbers[i];
      var htmlData = null;

      if (!existsSync(outputFileDirectory)) {
        mkdirSync(outputFileDirectory);
      }
      if (!existsSync(`${outputFileDirectory}htmlFiles/`)) {
        mkdirSync(`${outputFileDirectory}htmlFiles/`);
      }
      if (!existsSync(`${outputFileDirectory}htmlFiles/${seatNum['SeatNumber']}.html`)) {
        htmlData = await this.doRequestCall(`${baseurl}${seatNum['SeatNumber']}`);
        writeFileSync(`${outputFileDirectory}htmlFiles/${seatNum['SeatNumber']}.html`, htmlData);
      } else {
        htmlData = readFileSync(`${outputFileDirectory}htmlFiles/${seatNum['SeatNumber']}.html`);
      }

      // HTML Processing


      sheet.addRow({
        SeatNumber: seatNum['SeatNumber'],
        Name: '',
        ScienceMarks: 0,
        MathsMarks: 0,
        EnglishMarks: 0,
        GujaratiMarks: 0,
        SanskritMarks: 0,
      });
    }

    sheet.columns.forEach((column) => {
      column.width = column.header.length < 10 ? 10 : column.header.length;
    });
    sheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        bgColor: { argb: 'fdf731' },
        fgColor: { argb: 'fdf731' },
      };
    });

    res.attachment(
      `HSC-Sci-Result-${moment().format('YYYY-MM-DD hh:mm:ss a')}.xlsx`,
    );
    await workbook.xlsx.write(res);
  }


  // ======UTIL METHODS====== //
  async doRequestCall(url: string) {
    if(!url.startsWith('http://')) {
      url = `http://${url}`;
    }
    let response = this.httpService.get(url);
    try {
      let resp = await firstValueFrom(response);
      // console.log(url, 'success', resp.data);
      return resp.data;
    } catch (ex) {
      console.log(url, 'failure', ex);
      return null;
    }
  }

  async csvToRows(filename: string): Promise<any[]> {
    const results = <any>[];

    return new Promise((resolve, reject) => {

      fs.createReadStream(filename)
        .pipe(csv({}))
        .on('data', (data) => results.push(data))
        .on('end', () => {
          resolve(results);
        });
    });
  }
}
