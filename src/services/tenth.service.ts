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
export class TenthService {
  constructor(private httpService: HttpService) { }

  getHello(): string {
    return 'Hello World!';
  }

  async scrapRecurrsively(res: Response) {
    var servers = await this.csvToRows('files/servers.csv');
    console.log('servers', servers);

    for (let i = 0; i < servers.length; i++) {
      const server = servers[i];
      console.log(`[Scrapping from ${server['Servers']}] [${i}/${servers.length}]`);
      this.scrapResultsClass10(server['Servers'], res);
    }
  }

  async scrapResultsClass10(baseurl: string, res: Response) {
    var inputFilePath = 'files/input/gm-ssc.csv';
    var outputFileDirectory = `files/output${inputFilePath.substring(inputFilePath.lastIndexOf('/'), inputFilePath.lastIndexOf('.'))}/`;
    console.log('HAELLOEAO', inputFilePath, outputFileDirectory);

    var seatNumbers = await this.csvToRows(inputFilePath);
    console.log('seatNumbers', seatNumbers);

    const workbook = new Workbook();
    workbook.creator = 'Bhagwan';
    workbook.created = new Date();
    const sheet = workbook.addWorksheet('SSC-Results');

    let columns = [];
    let rowsData = [];

    // todo: api call
    for (let i = 0; i < seatNumbers.length; i++) {
      try {
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
          htmlData = await this.doRequestCall(`${baseurl}285soipmahc/ssc/${seatNum['SeatNumber'].substring(0, 3)}/${seatNum['SeatNumber'].substring(3, 5)}/${seatNum['SeatNumber']}.html`);
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

        const obtainedMarks = $('tr.background1 td[colspan="4"] b.texcolor').text().split(' (')[0];
        const overallGrade = $('tr.textcolor.background1.highlitedtd td').eq(0).text().trim();
        const percentileRank = $('tr.textcolor.background1.highlitedtd td').eq(1).text().trim();

        // Fetch the subject-wise details
        let totalMarks = 0;
        const subjectDetails = [];
        const subjectRows = $('tr:not(.textcolor):not(.background1):has(td[align])');
        subjectRows.map((index, element) => {
          if (element.children.length === 5 || element.children.length === 2) {
            let subjectName, marksExternal = '', marksInternal = '', marksTotal = '', subjectGrade = '';
            const row = $(element);
            subjectName = row.find('td').eq(0).text().trim();
            if (element.children.length === 5) {
              marksExternal = row.find('td').eq(1).text().trim();
              marksInternal = row.find('td').eq(2).text().trim();
              marksTotal = row.find('td').eq(3).text().trim();
              subjectGrade = row.find('td').eq(4).text().trim();
              totalMarks += 100;
            } else {
              subjectGrade = row.find('td').eq(1).text().trim();
            }

            subjectDetails.push({
              subject: subjectName,
              marksExternal: marksExternal,
              marksInternal: marksInternal,
              totalMarks: marksTotal,
              grade: subjectGrade
            });
          }

        }).get();

        subjectDetails.forEach((subject) => {
          if (!columns.find((column) => (column.header as string).includes(`${subject.subject}`))) {
            columns.push({ header: `${subject.subject}-ExternalMarks`, key: `${subject.subject}-ExternalMarks` });
            columns.push({ header: `${subject.subject}-InternalMarks`, key: `${subject.subject}-InternalMarks` });
            columns.push({ header: `${subject.subject}-TotalMarks`, key: `${subject.subject}-TotalMarks` });
            columns.push({ header: `${subject.subject}-Grade`, key: `${subject.subject}-Grade` });
          }
        });

        let data = {
          SeatNumber: seatNum['SeatNumber'],
          Name: name,
          SID: sid,
          Result: result,
          TotalMarks: totalMarks,
          ObtainedMarks: obtainedMarks,
          Grade: overallGrade,
          Percentile: percentileRank,
          Percentage: `${((parseFloat(obtainedMarks) / totalMarks) * 100).toFixed(2)}%`
        };

        subjectDetails.forEach((subject) => {
          data[`${subject.subject}-ExternalMarks`] = subject.marksExternal;
          data[`${subject.subject}-InternalMarks`] = subject.marksInternal;
          data[`${subject.subject}-TotalMarks`] = subject.totalMarks;
          data[`${subject.subject}-Grade`] = subject.grade;
        });
        rowsData.push(data);
      } catch (e) {
        console.log('Error', e);
      }
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

    // TEST : BETA ANALYSIS SHEET
    const analysisSheet = workbook.addWorksheet('SSC-Analysis');

    const resultPercentage = `${((rowsData.filter((row) => row.Result.includes('QUALIFIED')).length / rowsData.length) * 100).toFixed(2)}%`;
    let resultPercentageRow = analysisSheet.addRow(['Result Percentage', resultPercentage]);
    resultPercentageRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      bgColor: { argb: 'fffdf731' },
      fgColor: { argb: 'fffdf731' },
    };
    resultPercentageRow.font = { bold: true };

    /////
    analysisSheet.addRows(['', '']);
    /////

    let resultAnalysisHeaderRow = analysisSheet.addRow(['Result wise analysis']);
    resultAnalysisHeaderRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      bgColor: { argb: 'fffdf731' },
      fgColor: { argb: 'fffdf731' },
    };
    resultAnalysisHeaderRow.font = { bold: true };

    analysisSheet.addRow(['Result', 'Count']);

    const resultWiseAnalysis = {};
    rowsData.forEach((row) => {
      if (resultWiseAnalysis[row.Result]) {
        resultWiseAnalysis[row.Result] += 1;
      } else {
        resultWiseAnalysis[row.Result] = 1;
      }
    });

    Object.keys(resultWiseAnalysis).sort().forEach((key) => {
      analysisSheet.addRow([key, resultWiseAnalysis[key]]);
    });

    /////
    analysisSheet.addRows(['', '']);
    /////

    let gradeWiseAnalysisHeaderRow = analysisSheet.addRow(['Grade wise analysis']);
    gradeWiseAnalysisHeaderRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      bgColor: { argb: 'fffdf731' },
      fgColor: { argb: 'fffdf731' },
    };
    gradeWiseAnalysisHeaderRow.font = { bold: true };

    analysisSheet.addRow(['Grade', 'Count']);
    const gradeWiseAnalysis = {};
    rowsData.forEach((row) => {
      if (gradeWiseAnalysis[row.Grade]) {
        gradeWiseAnalysis[row.Grade] += 1;
      } else {
        gradeWiseAnalysis[row.Grade] = 1;
      }
    });
    Object.keys(gradeWiseAnalysis).sort().forEach((key) => {
      analysisSheet.addRow([key, gradeWiseAnalysis[key]]);
    });

    /////
    analysisSheet.addRows(['', '']);
    /////

    let subjectWiseAnalysisHeaderRow = analysisSheet.addRow(['Subject wise grade wise analysis']);
    subjectWiseAnalysisHeaderRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      bgColor: { argb: 'fffdf731' },
      fgColor: { argb: 'fffdf731' },
    };
    subjectWiseAnalysisHeaderRow.font = { bold: true };

    analysisSheet.addRow(['Subject', 'Grade', 'Count']);
    const subjectWiseGradeAnalysis = {};
    rowsData.forEach((row) => {
      Object.keys(row).forEach((key) => {
        if (key.includes('Grade') && key !== 'Grade') {
          if (subjectWiseGradeAnalysis[key]) {
            if (subjectWiseGradeAnalysis[key][row[key]]) {
              subjectWiseGradeAnalysis[key][row[key]] += 1;
            } else {
              subjectWiseGradeAnalysis[key][row[key]] = 1;
            }
          } else {
            subjectWiseGradeAnalysis[key] = {};
            subjectWiseGradeAnalysis[key][row[key]] = 1
          }
        }
      });
    });
    Object.keys(subjectWiseGradeAnalysis).forEach((key) => {
      Object.keys(subjectWiseGradeAnalysis[key]).sort().forEach((grade) => {
        analysisSheet.addRow([key, grade, subjectWiseGradeAnalysis[key][grade]]);
      });
    });

    analysisSheet.getColumn(1).width = 40;

    res.attachment(
      `SSC-GEN-Result-${moment().format('YYYY-MM-DD hh:mm:ss a')}.xlsx`,
    );
    await workbook.xlsx.write(res);
  }


  // ======UTIL METHODS====== //
  async doRequestCall(url: string) {

    console.log(env.NODE_TLS_REJECT_UNAUTHORIZED);
    if (!url.startsWith('http')) {
      url = `http://${url}`;
    }
    console.log(url);
    let response = this.httpService.get(url, {});
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
