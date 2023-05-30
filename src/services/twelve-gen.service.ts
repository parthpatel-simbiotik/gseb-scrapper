import { HttpService } from '@nestjs/axios';
import { Injectable } from '@nestjs/common';
import { Workbook } from 'exceljs';
import { Response } from 'express';
import { existsSync, mkdir, mkdirSync, readFileSync, writeFileSync } from 'fs';
import * as moment from 'moment';
import { env } from 'process';
import { firstValueFrom } from 'rxjs';
const fs = require('fs');
const cheerio = require('cheerio');
const csv = require('csv-parser');

@Injectable()
export class TwelveGenService {
  constructor(private httpService: HttpService) { }

  async scrapRecurrsively(res: Response) {
    var servers = await this.csvToRows('files/servers.csv');
    console.log('servers', servers);

    for (let i = 0; i < servers.length; i++) {
      const server = servers[i];
      console.log(`[Scrapping from ${server['Servers']}] [${i}/${servers.length}]`);
      this.scrapResultsClass12(server['Servers'], res);
    }
  }

  async scrapResultsClass12(baseurl: string, res: Response) {
    var inputFilePath = 'files/input/test1.csv';
    var outputFileDirectory = `files/output${inputFilePath.substring(inputFilePath.lastIndexOf('/'), inputFilePath.lastIndexOf('.'))}/`;
    console.log('HAELLOEAO', inputFilePath, outputFileDirectory);

    var seatNumbers = await this.csvToRows(inputFilePath);
    console.log('seatNumbers', seatNumbers);

    const workbook = new Workbook();
    workbook.creator = 'Bhagwan';
    workbook.created = new Date();
    const sheet = workbook.addWorksheet('HSC-GenResults');

    let columns = [];
    let rowsData = [];

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
        htmlData = await this.doRequestCall(`${baseurl}3015ecruosnepo/gen/${seatNum['SeatNumber'].substring(0, 3)}/${seatNum['SeatNumber'].substring(3, 5)}/${seatNum['SeatNumber']}.html`);
        if (htmlData != null && htmlData.includes('Please try again.')) {
          console.log('Please try again. received for ', seatNum['SeatNumber']);
          continue;
        }
        writeFileSync(`${outputFileDirectory}htmlFiles/${seatNum['SeatNumber']}.html`, htmlData);
      } else {
        htmlData = readFileSync(`${outputFileDirectory}htmlFiles/${seatNum['SeatNumber']}.html`);
      }

      let $ = cheerio.load(htmlData.toString());

      // Fetch the name
      const nameElement = $('b.textcolor:contains("Name:")');
      const name = nameElement.parent().text().trim().replace('Name:', '').trim();

      const sidElement = $('b.textcolor:contains("SID:")');
      const sid = sidElement.parent().text().trim().replace('SID:', '').trim();

      const resultElement = $('b.textcolor:contains("Result:")');
      const result = resultElement.parent().text().trim().replace('Result:', '').trim();

      const percentileElement = $('b.textcolor:contains("Percentile:")');
      const percentile = percentileElement.parent().text().trim().replace('Percentile:', '').trim();

      const gradeElement = $('b.textcolor:contains("Grade:")');
      const grade = gradeElement.parent().text().trim().replace('Grade:', '').trim();

      const totalMarksElement = $('td.textcolor:contains("Total Marks")');
      const totalMarks = totalMarksElement.parent().find('span').contents().eq(1).text();
      const obtainedMarks = totalMarksElement.parent().find('span').contents().eq(2).text();

      // Fetch the subject-wise details
      const subjectDetails = [];
      $('table.maintbl tr').each((index, row) => {
        const className = $(row).attr('class');
        if (className !== 'background1') {
          const columns = $(row).find('td span');
          const subject = $(columns[0]).text().trim();
          const totalMarks = $(columns[1]).text().trim();
          const obtainedMarks = $(columns[2]).text().trim();
          const grade = $(columns[3]).text().trim();

          subjectDetails.push({
            subject: subject,
            totalMarks: totalMarks,
            obtainedMarks: obtainedMarks,
            grade: grade
          });
        }
      });

      subjectDetails.forEach((subject) => {
        if (!columns.find((column) => (column.header as string).includes(`${subject.subject}`))) {
          // columns.push({ header: `${subject.subject}-TotalMarks`, key: `${subject.subject}-TotalMarks` });
          columns.push({ header: `${subject.subject}`, key: `${subject.subject}` });
          // columns.push({ header: `${subject.subject}-Grade`, key: `${subject.subject}-Grade` });
        }
      });

      let data = {
        SeatNumber: seatNum['SeatNumber'],
        Name: name,
        SID: sid,
        Result: result,
        TotalMarks: totalMarks,
        ObtainedMarks: obtainedMarks,
        Grade: grade,
        Percentile: `${percentile}%`,
        Percentage: `${((parseFloat(obtainedMarks) / parseFloat(totalMarks)) * 100).toFixed(2)}%`
      };

      subjectDetails.forEach((subject) => {
        // data[`${subject.subject}-TotalMarks`] = subject.totalMarks;
        data[`${subject.subject}`] = subject.obtainedMarks;
        // data[`${subject.subject}-Grade`] = subject.grade;
      });
      rowsData.push(data);
    }

    sheet.columns = [
      { header: 'SeatNumber', key: 'SeatNumber' },
      { header: 'Name', key: 'Name' },
      { header: 'SID', key: 'SID' },
      { header: 'Result', key: 'Result' },
      { header: 'TotalMarks', key: 'TotalMarks' },
      { header: 'ObtainedMarks', key: 'ObtainedMarks' },
      { header: 'Grade', key: 'Grade' },
      { header: 'Percentile', key: 'Percentile' },
      { header: 'Percentage', key: 'Percentage' },
      ...columns,
    ];

    sheet.addRows(rowsData);

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
      `HSC-Gen-Result-${moment().format('YYYY-MM-DD hh:mm:ss a')}.xlsx`,
    );
    await workbook.xlsx.write(res);
  }


  // ======UTIL METHODS====== //
  async doRequestCall(url: string) {

    if (!url.startsWith('http')) {
      url = `http://${url}`;
    }
    console.log(url);
    let response = this.httpService.get(url, {});
    try {
      let resp = await firstValueFrom(response);
      // console.log(url, 'success', resp.status, resp.statusText);
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