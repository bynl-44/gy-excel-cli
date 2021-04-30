import chalk from 'chalk';
import xlsx from 'xlsx';

export const SUCCESS = 'success';
export const INFO = 'info';
export const WARNING = 'warning';
export const ERROR = 'error';

const XLSX = 'xlsx';

export const MONTHS = [
  'C:1月',
  'D:2月',
  'E:3月',
  'F:4月',
  'G:5月',
  'H:6月',
  'I:7月',
  'J:8月',
  'K:9月',
  'L:10月',
  'M:11月',
  'N:12月',
];

/**
 * 格式化消息输出
 *
 * @param {string} [status=INFO] 状态标志，默认为: INFO
 * @param {string} message 消息内容
 * @returns
 */
export const message = (status = INFO, message) => {
  if (status === INFO) {
    return `[${status}] ${message}`;
  }
  if (status === WARNING) {
    return chalk.yellow(`[${status}] ${message}`);
  }
  if (status === ERROR) {
    return chalk.red(`[${status}] ${message}`);
  }
  if (status === SUCCESS) {
    return chalk.green(`[${status}] ${message}`);
  }
};

export const isExcel = (path) => {
  const suffix = path.substring(path.lastIndexOf('.') + 1);
  return suffix === XLSX;
};

/**
 * 给数据进行排序，排序形式：[A1,B1,C1,A2,B2,C2]
 *
 * @param {Array} cells
 * @returns {Array}
 */
export const cellSort = (cells) => {
  return Object.keys(cells)
    .filter((cell) => {
      // 获取单元格编号（如：第一个单元格为A1)
      const cellReg = /^[A-Z]\d{1,2}$/g;
      return cell.match(cellReg);
    })
    .sort((a, b) => {
      // 第一次按列排序,即（A，B，C，D）
      const columnA = a.substring(0, 1).charCodeAt();
      const columnB = b.substring(0, 1).charCodeAt();

      return columnA - columnB;
    })
    .sort((a, b) => {
      // 第二次按行排序,即（1，2，3，4）
      const rowA = parseInt(a.substring(1));
      const rowB = parseInt(b.substring(1));

      return rowA - rowB;
    });
};

/**
 * 从读取出的数据中，提取有用的数据，并进行排序、封装转换为格式化数组
 *
 * @param {xlsx.Sheet} sheet
 * @returns
 */
export const formatData = (sheet) => {
  const _array = [];
  const cells = cellSort(sheet);
  // 获取行号
  const uniqueRowIds = new Set([...cells.map((v) => parseInt(v.substring(1)))]);

  [...uniqueRowIds].map((id) => {
    // 取出每一行数据
    if (id !== 1) {
      // 去除第一行表头的标题
      let person = {};
      cells.map((cell, index) => {
        // 取出该行每一个单元格里的数据
        const row = parseInt(cell.substring(1));
        if (id === row) {
          person[cell] = sheet[cell].v;
        }
      });

      _array.push(person);
    }
  });

  return _array;
};

export function errorHandler(err) {
  const pathReg = /TypeError: path/;
  const code = err.code || '';
  const msg = err.message;
  if (code === 'path' || pathReg.test(err)) {
    console.log(message(ERROR, msg || `文件路径错误，请输入正确的文件路径`));
  }
  console.log(err);
}

export default {};
