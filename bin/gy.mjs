#!/usr/bin/env node

import { Command } from 'commander';
import inquirer from 'inquirer';
import {
  checkFiles,
  analyzeFiles,
  calculate,
  writeFile,
} from '../src/file.mjs';
import { message, INFO, SUCCESS, MONTHS, errorHandler } from '../src/util.mjs';
import moment from 'moment';
import _ from 'lodash';

const { isEmpty } = _;

const program = new Command();

program.version('1.0.0');

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
