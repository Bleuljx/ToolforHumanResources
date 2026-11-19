/* ============================================================
 * zip.js — 极简 ZIP 读写（浏览器/Node 通用）
 * 写入：优先 deflate-raw 压缩，不支持则 store；文件名 UTF-8（flag 0x0800）
 * ============================================================ */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ZipUtil = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  'use strict';

  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(u8) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function utf8(str) {
    return new TextEncoder().encode(str);
  }

  var CAN_COMPRESS = (function () {
    try { return typeof CompressionStream !== 'undefined' && typeof Blob !== 'undefined'; }
    catch (e) { return false; }
  })();

  function deflateRaw(u8) {
    if (!CAN_COMPRESS) return Promise.resolve(null);
    try {
      var stream = new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate-raw'));
      return new Response(stream).arrayBuffer().then(function (ab) {
        return new Uint8Array(ab);
      }).catch(function () { return null; });
    } catch (e) {
      return Promise.resolve(null);
    }
  }

  function u16(v) { return [v & 0xFF, (v >>> 8) & 0xFF]; }
  function u32(v) { return [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF]; }

  /** files: [{name:'a/b.xml', data:Uint8Array|string}] -> Promise<Uint8Array>
   *  onProgress(done, total)：每完成一个文件的压缩回调一次（可选） */
  function zip(files, onProgress) {
    var entries = [];
    var chain = Promise.resolve();
    files.forEach(function (f, fi) {
      chain = chain.then(function () {
        var raw = typeof f.data === 'string' ? utf8(f.data) : f.data;
        return deflateRaw(raw).then(function (comp) {
          var body = comp && comp.length < raw.length ? comp : raw;
          entries.push({
            nameBytes: utf8(f.name),
            data: body,
            crc: crc32(raw),
            rawSize: raw.length,
            compSize: body.length,
            method: body === raw ? 0 : 8
          });
          if (onProgress) { try { onProgress(fi + 1, files.length); } catch (e) { } }
        });
      });
    });

    return chain.then(function () {
      var chunks = [];
      var central = [];
      var offset = 0;
      var now = new Date();
      var dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() / 2)) & 0xFFFF;
      var dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;

      entries.forEach(function (e) {
        var local = [].concat(
          u32(0x04034b50), u16(20), u16(0x0800), u16(e.method),
          u16(dosTime), u16(dosDate),
          u32(e.crc), u32(e.compSize), u32(e.rawSize),
          u16(e.nameBytes.length), u16(0)
        );
        var header = new Uint8Array(local);
        chunks.push(header, e.nameBytes, e.data);
        var c = [].concat(
          u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(e.method),
          u16(dosTime), u16(dosDate),
          u32(e.crc), u32(e.compSize), u32(e.rawSize),
          u16(e.nameBytes.length), u16(0), u16(0), u16(0), u16(0),
          u32(0), u32(offset)
        );
        central.push(new Uint8Array(c), e.nameBytes);
        offset += header.length + e.nameBytes.length + e.data.length;
      });

      var centralStart = offset;
      var centralSize = 0;
      central.forEach(function (a) { centralSize += a.length; });

      var eocd = new Uint8Array([].concat(
        u32(0x06054b50), u16(0), u16(0),
        u16(entries.length), u16(entries.length),
        u32(centralSize), u32(centralStart), u16(0)
      ));

      var all = chunks.concat(central, [eocd]);
      var total = all.reduce(function (s, a) { return s + a.length; }, 0);
      var out = new Uint8Array(total);
      var p = 0;
      all.forEach(function (a) { out.set(a, p); p += a.length; });
      return out;
    });
  }

  /* ---------------- 解包读取 ---------------- */
  function rd16(u8, p) { return u8[p] | (u8[p + 1] << 8); }
  function rd32(u8, p) { return (u8[p] | (u8[p + 1] << 8) | (u8[p + 2] << 16) | (u8[p + 3] << 24)) >>> 0; }

  function findEocd(u8) {
    var end = u8.length, lo = Math.max(0, end - 22 - 65535);
    for (var i = end - 22; i >= lo; i--) {
      if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x05 && u8[i + 3] === 0x06) return i;
    }
    return -1;
  }

  function inflateRaw(data) {
    if (typeof DecompressionStream === 'undefined') {
      return Promise.reject(new Error('当前环境不支持解压 deflate 压缩包，请改用「仅存储」方式的压缩包'));
    }
    try {
      var stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new Response(stream).arrayBuffer().then(function (ab) { return new Uint8Array(ab); });
    } catch (e) { return Promise.reject(e); }
  }

  /** u8:Uint8Array -> Promise<{name:string, data:Uint8Array}[]>，跳过目录项 */
  function unzip(u8) {
    var eocd = findEocd(u8);
    if (eocd < 0) return Promise.reject(new Error('不是有效的 zip 文件'));
    var n = rd16(u8, eocd + 10), cdOff = rd32(u8, eocd + 16), p = cdOff;
    var entries = [];
    for (var i = 0; i < n; i++) {
      if (p + 46 > u8.length || rd32(u8, p) !== 0x02014b50) return Promise.reject(new Error('zip 目录损坏'));
      var method = rd16(u8, p + 10), csize = rd32(u8, p + 20),
        nameLen = rd16(u8, p + 28), extraLen = rd16(u8, p + 30), cmtLen = rd16(u8, p + 32),
        lho = rd32(u8, p + 42);
      var name = new TextDecoder('utf-8').decode(u8.subarray(p + 46, p + 46 + nameLen));
      p += 46 + nameLen + extraLen + cmtLen;
      if (name.slice(-1) === '/' || /(^|\/)(__MACOSX|\.DS_Store)/i.test(name)) continue;
      var lnLen = rd16(u8, lho + 26), leLen = rd16(u8, lho + 28);
      var start = lho + 30 + lnLen + leLen;
      entries.push({ name: name, method: method, data: u8.subarray(start, start + csize) });
    }
    return Promise.all(entries.map(function (e) {
      if (e.method === 0) return Promise.resolve({ name: e.name, data: new Uint8Array(e.data) });
      return inflateRaw(e.data).then(function (d) { return { name: e.name, data: d }; });
    }));
  }

  return { zip: zip, unzip: unzip, crc32: crc32, utf8: utf8 };
});
