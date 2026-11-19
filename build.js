/* build.js — 把所有 JS/数据内联成单个 HTML，双击即可离线使用 */
const fs = require('fs');
const path = require('path');
const root = __dirname;

function read(p) { return fs.readFileSync(path.join(root, p), 'utf8'); }
function js(p) { return read(p).replace(/<\/script/gi, '<\\/script'); }

const html = read('src/index.html');
const opts = JSON.parse(read('src/options.json'));

const parts = [
  js('vendor/xlsx.js'),
  js('src/zip.js'),
  js('src/xlsxw.js'),
  js('src/core.js'),
  js('src/att-core.js'),
  js('src/idle-core.js'),
  js('src/ren-core.js'),
  'window.SALARY_OPTIONS=' + JSON.stringify(opts) + ';',
  js('src/app.js'),
  js('src/att-app.js'),
  js('src/idle-app.js'),
  js('src/ren-app.js')
];

const block = parts.map(function (s) { return '<script>\n' + s + '\n</script>'; }).join('\n');
// 注意：必须用函数式替换，否则字符串里的 $' $& 等会被当成替换模式
const out = html.replace('</body>', function () { return block + '\n</body>'; });

if (out === html) { console.error('未找到 </body> 插入点'); process.exit(1); }

const dist = path.join(root, 'dist');
if (!fs.existsSync(dist)) fs.mkdirSync(dist);
const target = path.join(dist, '人力工具集成平台.html');
fs.writeFileSync(target, out, 'utf8');
console.log('built:', target, (out.length / 1024).toFixed(0) + ' KB');
