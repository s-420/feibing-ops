const assert = require('assert');
const { findExcelUrl } = require('./渠道欢迎语-批量导入-Excel');

assert.strictEqual(
  findExcelUrl({ code: 0, data: { url: 'https://example.com/file.xlsx' } }),
  'https://example.com/file.xlsx',
);
assert.strictEqual(findExcelUrl({ data: { url: 'https://example.com/file.png' } }), '');
assert.strictEqual(
  findExcelUrl([{ nested: { downloadUrl: 'https://example.com/b.xlsx?token=redacted' } }]),
  'https://example.com/b.xlsx?token=redacted',
);

console.log('渠道欢迎语 Excel 批量导入脚本测试通过');
