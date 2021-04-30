import lodash from 'lodash';
import inquirer from 'inquirer';
import fs from 'fs';
import xlsx from 'xlsx';
import { message, isExcel, INFO, SUCCESS } from './util.mjs';
import { formatData } from './util.mjs';
import { strip, times } from 'number-precision';

const { isEmpty } = lodash;

/**
 * 读取文件数据
 *
 * @param {string} path
 * @returns {xlsx.WorkBook} workbook
 */
function readFile(path) {
  if (!isExcel(path)) {
    const err = new Error();
    err.code = 'path';
    err.message = '文件类型错误';
    throw err;
  }
  const buf = fs.readFileSync(path);
  const workbook = xlsx.read(buf, { type: 'buffer' });

  return workbook;
}

/**
 * 分析from文件数据,把数据处理为容易处理的形式
 *
 * @param {xlsx.WorkBook} workbook
 * @returns
 */
function analyzeFromData(workbook) {
  const { SheetNames, Sheets } = workbook;
  const sheet1 = Sheets[SheetNames[0]];
  const sheet2 = Sheets[SheetNames[1]];

  const personalPerformances = formatData(sheet1);
  const personalOutputs = formatData(sheet2);

  return { personalPerformances, personalOutputs };
}

function analyzeToData(workbook) {
  const { SheetNames, Sheets } = workbook;
  const sheet1 = Sheets[SheetNames[0]];

  return formatData(sheet1);
}

/**
 * 如果文件路径为空，就提醒输入文件路径
 *
 * @param {string} [name='from'] 哪个文件（from/to）
 * @param {string} [message=''] 提示信息
 * @param {string} path 文件路径
 * @returns {string} path 文件路径
 */
async function emptyPath(name = 'from', message = '', path) {
  if (isEmpty(path)) {
    path = await inquirer.prompt([
      {
        type: 'input',
        name: name,
        message: message,
      },
    ]);
    if (isEmpty(path[name])) {
      const err = new Error();
      err.code = 'path';
      err.message = '文件路径不能为空';
      throw err;
    }
  }

  return path;
}

/**
 * 检查[from]文件
 *
 * @param {string} from 文件路径
 * @returns
 */
async function checkFromFile(from) {
  from = await emptyPath('from', '请输入要转换的Excel文件路径:', from);

  console.log(message(INFO, '开始检查From文件数据...'));
  return readFile(from);
}

async function checkToFile(to) {
  to = await emptyPath('to', '请输入目标Excel文件路径:', to);

  return readFile(to);
}

export async function checkFiles(from, to) {
  const fromWB = await checkFromFile(from);
  console.log(message(SUCCESS, 'From文件数据检查完成！'));
  console.log(message(INFO, '开始检查To文件数据...'));
  const toWB = await checkToFile(to);
  console.log(message(SUCCESS, 'To文件数据检查完成！'));

  return { fromWB, toWB };
}

export function analyzeFiles(from, to) {
  const fromSheets = analyzeFromData(from);
  const toSheet = analyzeToData(to);

  return { fromSheets, toSheet };
}

export function calculate(from, to, month) {
  const { personalPerformances, personalOutputs } = from;
  const column = month.substring(0, 1);

  const calculatedTo = to.map((row, index) => {
    const jobNo = row[`C${index + 2}`];
    const monthCell = `E${index + 2}`;
    const shouldCell = `F${index + 2}`;

    if (jobNo === '工号') {
      row[monthCell] = `${month.substring(month.indexOf(':') + 1)}绩效`;
    } else if (!isEmpty(jobNo)) {
      personalOutputs.map((output, i) => {
        const outputJobNo = output[`A${i + 2}`];
        if (jobNo === outputJobNo) {
          row[monthCell] = times(
            parseFloat(output[`${column}${i + 2}`] || 0),
            0.2,
            10000
          );
        }
      });

      personalPerformances.map((performance, i) => {
        const performanceNo = performance[`A${i + 2}`];
        if (jobNo === performanceNo) {
          row[shouldCell] = strip(parseFloat(performance[`Q${i + 2}`]));
        }
      });
    }

    return row;
  });

  return calculatedTo;
}

export function writeFile(toSheet, toWB, fileName = '矿山院月终.xlsx') {
  let sheet = toWB.Sheets[toWB.SheetNames[0]];
  sheet = { ...sheet, ...toSheet };
  toWB.Sheets[toWB.SheetNames[0]] = sheet;

  xlsx.writeFile(toWB, fileName);
}

export default {};
