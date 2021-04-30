#!/usr/bin/env node

const { Command } = require('commander');
const inquirer = require('inquirer');
const {
  checkFiles,
  analyzeFiles,
  calculate,
  writeFile,
} = require('../src/file.js');
const {
  INFO,
  SUCCESS,
  MONTHS,
  errorHandler,
  message,
} = require('../src/util.js');
const moment = require('moment');
const _ = require('lodash');
const pkg = require('../package.json');

const { isEmpty } = _;

const program = new Command();

program.version(pkg.version);

program
  .arguments('[from] [to]')
  .description('Displays help information.', {
    from: '源文件',
    to: '目标文件',
  })
  .action(async (from, to) => {
    try {
      console.log(`Start...`);
      // check files
      if (!isEmpty(from) && !isEmpty(to))
        console.log(message(INFO, `开始检查输入数据完整性...`));
      const { fromWB, toWB } = await checkFiles(from, to);
      console.log(message(SUCCESS, `数据检查完成！`));

      // analyze file content
      console.log(message(INFO, `开始分析输入数据...`));
      const { fromSheets, toSheet } = await analyzeFiles(fromWB, toWB);
      console.log(message(SUCCESS, `数据分析完成！`));

      const { month } = await inquirer.prompt([
        {
          type: 'list',
          name: 'month',
          message: '选择月份：',
          choices: MONTHS,
          default: () =>
            `${String.fromCharCode('B'.charCodeAt() + moment().month() + 1)}:${
              moment().month() + 1
            }月`,
        },
      ]);
      // calculate
      console.log(message(INFO, `开始计算数据...`));
      const calculatedToSheet = calculate(fromSheets, toSheet, month);
      console.log(message(SUCCESS, `数据计算完成！`));

      // write file
      console.log(message(INFO, `开始写入数据...`));
      writeFile(calculatedToSheet, toWB);
      console.log(message(SUCCESS, `数据写入完成！`));
    } catch (e) {
      errorHandler(e);
    }
  });

program.parse(process.argv);
