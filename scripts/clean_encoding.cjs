// 清理文件中的不可见字符（保留中文和英文）
const fs = require('fs');

const filePath = process.argv[2];
if (!filePath) {
  console.log('Usage: node scripts/clean_encoding.cjs <file>');
  process.exit(1);
}

const content = fs.readFileSync(filePath, 'utf8');
const originalLen = content.length;

// 移除不在保留范围内的字符
// 保留: ASCII (0x00-0x7F), 中文 (0x4E00-0x9FFF), 中文标点 (0x3000-0x303F), 全角字符 (0xFF00-0xFFEF)
const cleaned = content.replace(/[^\x00-\x7F\u4E00-\u9FFF\u3000-\u303F\uFF00-\uFFEF]/g, '');

const removedCount = originalLen - cleaned.length;
if (removedCount > 0) {
  fs.writeFileSync(filePath, cleaned, 'utf8');
  console.log(`Cleaned ${removedCount} invisible characters from ${filePath}`);
} else {
  console.log('No invisible characters found');
}
