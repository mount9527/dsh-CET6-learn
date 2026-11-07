/**
 * 极简 .docx 读取：只做一件事——把 word/document.xml 里的段落按行吐出来。
 *
 * 为什么不引第三方库：
 *   构建脚本要在使用者的机器上跑，多一个依赖就多一次「装不上」的机会。.docx 就是个
 *   zip 包，Node 自带的 zlib 足以解开 deflate 段，真正的代码只有下面这几十行。
 *
 * 为什么优先用 docx 而不是 PDF：
 *   六级的整卷 PDF 里，选词填空的空号（26–35）印在页边、被抽成独立的数字行，段落里的
 *   空位只剩下一个空白，无法还原；docx 里是行内的 `__26__`，段落也一段一行。
 *   同理，PDF 的页脚（站点水印）会混进正文，docx 没有。
 */

import fs from 'node:fs'
import zlib from 'node:zlib'

const EOCD_SIG = 0x06054b50
const CEN_SIG = 0x02014b50

/** 从 zip 里取一个文件（不落盘） */
function readZipEntry(file, wantName) {
  const buf = fs.readFileSync(file)
  // 从尾部往前找 EOCD（注释最长 65535 字节）
  let eocd = -1
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 66000); i--) {
    if (buf.readUInt32LE(i) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  if (eocd === -1) throw new Error('不是合法的 docx/zip：找不到中央目录')
  const count = buf.readUInt16LE(eocd + 10)
  let off = buf.readUInt32LE(eocd + 16)

  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(off) !== CEN_SIG) break
    const method = buf.readUInt16LE(off + 10)
    const compSize = buf.readUInt32LE(off + 20)
    const nameLen = buf.readUInt16LE(off + 28)
    const extraLen = buf.readUInt16LE(off + 30)
    const commentLen = buf.readUInt16LE(off + 32)
    const localOff = buf.readUInt32LE(off + 42)
    const name = buf.toString('utf8', off + 46, off + 46 + nameLen)
    if (name === wantName) {
      const lNameLen = buf.readUInt16LE(localOff + 26)
      const lExtraLen = buf.readUInt16LE(localOff + 28)
      const start = localOff + 30 + lNameLen + lExtraLen
      const data = buf.subarray(start, start + compSize)
      if (method === 0) return data
      if (method === 8) return zlib.inflateRawSync(data)
      throw new Error(`docx 用了不支持的压缩方式 ${method}`)
    }
    off += 46 + nameLen + extraLen + commentLen
  }
  throw new Error(`docx 里没有 ${wantName}`)
}

const ENTITIES = {
  '&lt;': '<',
  '&gt;': '>',
  '&amp;': '&',
  '&quot;': '"',
  '&apos;': "'",
  '&nbsp;': ' ',
}

const decode = (s) =>
  s
    .replace(/&(?:lt|gt|amp|quot|apos|nbsp);/g, (m) => ENTITIES[m])
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))

/**
 * 读出段落数组（空段落已丢掉）。
 * Word 会把一句话按排版拆成多个 run，这里按 <w:t> 拼接；<w:br/> 与段落等价处理。
 */
export function docxParagraphs(file) {
  const xml = readZipEntry(file, 'word/document.xml').toString('utf8')
  return xml
    .split(/<\/w:p>/)
    .map((chunk) =>
      decode(
        chunk
          .replace(/<w:tab\b[^>]*\/>/g, ' ')
          .replace(/<w:br\b[^>]*\/>/g, '\n')
          .replace(/<w:cr\b[^>]*\/>/g, '\n')
          .replace(/<[^>]+>/g, '')
      ).trim()
    )
    .flatMap((p) => p.split('\n'))
    .map((p) => p.replace(/\s+/g, ' ').trim())
    .filter(Boolean)
}

/** 只判断能不能读（构建开头用来决定用 docx 还是 pdf） */
export function isDocx(file) {
  return /\.docx$/i.test(file) && fs.existsSync(file)
}
