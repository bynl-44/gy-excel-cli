const CliTest = require('command-line-test');
const path = require('path');
const assert = require('assert');
const pkg = require('../package.json');

const binFile = path.resolve(pkg.bin.gy);

describe('gy', function () {
  it('`gy --version` should be ok', function () {
    const cliTest = new CliTest();
    return cliTest.execFile(binFile, ['--version'], {}).then((res) => {
      assert.strictEqual(res.stdout, pkg.version);
    });
  });

  it('`gy from to` should be reject', function () {
    const cliTest = new CliTest();
    return cliTest.execFile(binFile, ['from', 'to'], {}).then((res) => {
      assert.match(res.stdout, /文件类型错误/);
    });
  });
});
